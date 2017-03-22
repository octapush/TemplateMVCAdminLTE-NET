using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Ajax.Utilities;

namespace TemplateMVCAdminLTE.Core
{
    public static class HtmlHelperExtensions
    {
        public static MvcHtmlString JsMinify(this HtmlHelper htmlHelper,Func<object,object> markup )
        {
            var notMinifiedJs = (markup.DynamicInvoke(htmlHelper.ViewContext) ?? "").ToString();
            var minifier = new Minifier();
            var minifiedJs = minifier.MinifyJavaScript(notMinifiedJs, new CodeSettings
            {
                EvalTreatment = EvalTreatment.MakeImmediateSafe,
                PreserveImportantComments = false
            });
            return new MvcHtmlString(minifiedJs);
        }
    }
}