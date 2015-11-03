using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;

namespace ExcelMacroAndPDFExport
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var currentFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var sourceWorkbookFile = currentFolder + @"\Workbook.xlsx";
            var pdfOutputFile = currentFolder + @"\Workbook.pdf";
            var macroFile = currentFolder + @"\TestMacro.macro";
            var macroName = "TestMacro";

            try
            {
                // if the output file exsts then delete it
                if (File.Exists(pdfOutputFile))
                {
                    File.Delete(pdfOutputFile);
                }

                // get the excel instance, and wrap it in a lifetime manager to make sure it gets closed and cleaned from memory
                using (var excelAppLifetime = new ExcelApplicationLifetime(new Microsoft.Office.Interop.Excel.Application()))
                {
                    var excelApp = excelAppLifetime.Instance;
                    excelApp.ScreenUpdating = false;
                    excelApp.DisplayAlerts = false;
                    excelApp.Visible = false;

                    // get the workbook, and wrap it in a lifetime manager to make sure it gets closed and cleaned from memory
                    using (var workbookLifetime = new ExcelWorkbookLifetime(excelApp.Workbooks.Open(sourceWorkbookFile)))
                    {
                        var workbook = workbookLifetime.Instance;

                        // import the macro file
                        var module = workbook.VBProject.VBComponents.Import(macroFile);

                        // create the wierd string that excel needs to have to run this specific macro
                        var macro = string.Format("{0}!{1}.{2}", workbook.Name, module.Name, macroName);

                        // run the macro
                        excelApp.Run(macro);

                        // add a page footer
                        using (var worksheetLifetime = ComObjectLifetime.Create(workbook.Worksheets[1] as Worksheet))
                        {
                            var worksheet = worksheetLifetime.Instance;
                            worksheet.PageSetup.CenterFooter = "Page &P of " + worksheet.PageSetup.Pages.Count;
                        }

                        // export the result as PDF to PDF output file location
                        workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfOutputFile, OpenAfterPublish: true);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

            // Tell the garbage collector to collect any memory that's pending being cleaned up
            GC.Collect();
        }
    }

    public class ExcelApplicationLifetime : ComObjectLifetime<Microsoft.Office.Interop.Excel.Application>, IDisposable
    {
        public ExcelApplicationLifetime(Microsoft.Office.Interop.Excel.Application instance)
            : base(instance)
        {
        }

        public void Dispose()
        {
            this.Instance.Quit();
            base.Dispose();
        }
    }

    public class ExcelWorkbookLifetime : ComObjectLifetime<Microsoft.Office.Interop.Excel.Workbook>, IDisposable
    {
        public ExcelWorkbookLifetime(Microsoft.Office.Interop.Excel.Workbook instance)
            : base(instance)
        {
        }

        public void Dispose()
        {
            this.Instance.Close();
            base.Dispose();
        }
    }

    public class ComObjectLifetime<T> : IDisposable
    {

        public readonly T Instance;

        public ComObjectLifetime(T instance)
        {
            this.Instance = instance;
        }

        public virtual void Dispose()
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.Instance);
        }
    }

    public class ComObjectLifetime
    {
        public static ComObjectLifetime<T> Create<T>(T instance)
        {
            return new ComObjectLifetime<T>(instance);
        }
    }
}