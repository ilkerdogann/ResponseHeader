using System;
using System.IO;
using System.Web.Mvc;

namespace Abis.Web.Core.ActionResults
{
    public class WordResult : ActionResult
    {
        public WordResult(string viewName)
        {
            ViewName = viewName;
        }

        public string FileName { get; set; }
        public string ViewName { get; set; }
        public object Model { get; set; }

        public override void ExecuteResult(ControllerContext context)
        {
            context.Controller.ViewData.Model = Model;

            using (var stream = new MemoryStream())
            using (var writer = new StreamWriter(stream))
            {
                var viewResult = ViewEngines.Engines.FindPartialView(context, ViewName);
                var viewContext = new ViewContext(context, viewResult.View, context.Controller.ViewData, context.Controller.TempData, writer);
                viewResult.View.Render(viewContext, writer);
                viewResult.ViewEngine.ReleaseView(context, viewResult.View);
                writer.Flush();

                context.HttpContext.Response.Clear();
                context.HttpContext.Response.BinaryWrite(stream.ToArray());
                context.HttpContext.Response.ContentType = "application/msword";
                context.HttpContext.Response.AddHeader("content-disposition", $"filename={FileName}.doc");
                context.HttpContext.Response.End();
            }
        }
    }
}