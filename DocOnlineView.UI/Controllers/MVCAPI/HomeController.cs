using Aspose.Cells;
using Aspose.Slides.Pptx;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web.Http;

namespace DocOnlineView.UI.Controllers.MVCAPI
{
    public class HomeController : ApiController
    {
        [HttpGet]
        public DataTable CourseViewOnLine(string fileName)
        {
            DataTable dtlist = new DataTable();
            dtlist.Columns.Add("TempDocHtml", typeof(string));

            string fileDire = "/Files";//对于这种pfd如要嵌套在html中  所以走相对路径
            string sourceDoc = Path.Combine(fileDire, fileName);
            string saveDoc = "";

            string docExtendName = System.IO.Path.GetExtension(sourceDoc).ToLower();
            bool result = false;
            if (docExtendName == ".pdf")
            {
                //pdf模板文件，对于pdf需要先建立个temppdf.html把对应的js都引入到里面，，然后 后端生成html嵌套住pdf文件，，最后 js生成pdf显示
                string tempFile = Path.Combine(fileDire, "temppdf.html");
                saveDoc = Path.Combine(fileDire, "viewFiles/onlinepdf.html");//保存后的html
                result = PdfToHtml(
                      sourceDoc,
                       System.Web.HttpContext.Current.Server.MapPath(tempFile),
                      System.Web.HttpContext.Current.Server.MapPath(saveDoc));
            }
            else
            {
                saveDoc = Path.Combine(fileDire, "viewFiles/onlineview.html");
                result = OfficeDocumentToHtml(
                      System.Web.HttpContext.Current.Server.MapPath(sourceDoc),
                      System.Web.HttpContext.Current.Server.MapPath(saveDoc));
            }

            if (result)
            {
                dtlist.Rows.Add(saveDoc);
            }

            return dtlist;
        }

        private bool OfficeDocumentToHtml(string sourceDoc, string saveDoc)
        {
            bool result = false;

            //获取文件扩展名
            string docExtendName = System.IO.Path.GetExtension(sourceDoc).ToLower();
            switch (docExtendName)
            {
                case ".doc":
                case ".docx":
                    Aspose.Words.Document doc = new Aspose.Words.Document(sourceDoc);
                    doc.Save(saveDoc, Aspose.Words.SaveFormat.Html);

                    result = true;
                    break;
                case ".xls":
                case ".xlsx":
                    Workbook workbook = new Workbook(sourceDoc);
                    workbook.Save(saveDoc, SaveFormat.Html);

                    result = true;
                    break;
                case ".ppt":
                case ".pptx":
                    //templateFile = templateFile.Replace("/", "\\");
                    //string templateFile = sourceDoc;
                    //templateFile = templateFile.Replace("/", "\\");
                    PresentationEx pres = new PresentationEx(sourceDoc);
                    pres.Save(saveDoc, Aspose.Slides.Export.SaveFormat.Html);

                    result = true;
                    break;
                default:
                    break;
            }

            return result;
        }

        private bool PdfToHtml(string fileName, string tempFile, string saveDoc)
        {
            //---------------------读html模板页面到stringbuilder对象里---- 
            StringBuilder htmltext = new StringBuilder();
            using (StreamReader sr = new StreamReader(tempFile)) //模板页路径
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    htmltext.Append(line);
                }
                sr.Close();
            }

            fileName = fileName.Replace("\\", "/");
            //----------替换htm里的标记为你想加的内容 
            htmltext.Replace("$PDFFILEPATH", fileName);

            //----------生成htm文件------------------―― 
            using (StreamWriter sw = new StreamWriter(saveDoc, false,
                System.Text.Encoding.GetEncoding("utf-8"))) //保存地址
            {
                sw.WriteLine(htmltext);
                sw.Flush();
                sw.Close();

            }

            return true;
        }
    }
}
