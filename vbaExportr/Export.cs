using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;

namespace vbaExportr
{
    
    public static class Export
    {
        public class Reference
        {
            public bool isBuiltIn { get; set; }
            public string description { get; set; }
            public string fullPath { get; set; }
            public string guid { get; set; }
            public bool isBroken { get; set; }
            public int major { get; set; }
            public int minor { get; set; }
            public string name { get; set; }
        }
        public class ResponseObject
        {
            public string projectName { get; set; }
            public List<Reference> references { get; set; }
            public bool isProtected { get; set; }
            public List<BasFile> basFiles { get; set; }
        }
        public class BasFile
        {
            public string componentName { get; set; }
            public string code { get; set; }
        }
        
        public static List<ResponseObject> Open(string fileName)
        {
            List<ResponseObject> responseObjects = new List<ResponseObject>();
            var excel = new Excel.Application();
            var excelStillOpenLame = System.Diagnostics.Process.GetProcessesByName("EXCEL")[0];
            excel.DisplayAlerts = false;
            Console.WriteLine("Opening Workbook...");
            var workbook = excel.Workbooks.Open(fileName, false, true, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, false, true, Type.Missing);

            int projectCount = excel.VBE.VBProjects.Count;
            for (int i = 1; i <= projectCount; i++)
            {
                Console.WriteLine("Inspecting VBProject...");
                ResponseObject responseObj = new ResponseObject();
                List<BasFile> files = new List<BasFile>();
                //var project = workbook.VBProject;
                //var projects = project.Collection.VBE.VBProjects;
                var project = excel.VBE.VBProjects.Item(i);
                var projectName = project.Name;
                responseObj.projectName = projectName;

                List<Reference> references = new List<Reference>();
                //get all references in app
                for (int q = 1; q <= project.References.Count; q++)
                {
                    var reference = project.References.Item(q);
                    references.Add(new Reference()
                    {
                        isBuiltIn = reference.BuiltIn,
                        description = reference.Description,
                        fullPath = reference.FullPath,
                        guid = reference.Guid,
                        isBroken = reference.IsBroken,
                        major = reference.Major,
                        minor = reference.Minor,
                        name = reference.Name
                    });
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(reference);
                    reference = null;
                }
                responseObj.references = references;
                //get all code modules and mark instead if it was protected and this couldn't be accomplished.
                if (project.Protection == VBA.vbext_ProjectProtection.vbext_pp_locked)
                {
                    responseObj.isProtected = true;
                }
                else
                {
                    foreach (var component in project.VBComponents)
                    {
                        VBA.VBComponent vbComponent = component as VBA.VBComponent;
                        if (vbComponent != null)
                        {
                            string componentName = vbComponent.Name;
                            var basFile = new BasFile();

                            basFile.componentName = componentName;

                            var componentCode = vbComponent.CodeModule;
                            int componentCodeLines = componentCode.CountOfLines;

                            int line = 1;

                            string moduleCode = "'No code commited to file.";

                            if (componentCodeLines > 0)
                            {
                                moduleCode = componentCode.get_Lines(line, componentCodeLines);
                            }
                            basFile.code = moduleCode;
                            files.Add(basFile);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(vbComponent);
                            vbComponent = null;
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(componentCode);
                            componentCode = null;
                        }
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(project);
                project = null;
                responseObj.basFiles = files;
                responseObjects.Add(responseObj);
            }
            
            workbook.Close(false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            workbook = null;

            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            excel = null;
            if (!excelStillOpenLame.CloseMainWindow())
            {
                excelStillOpenLame.Kill();
            }
            return responseObjects;
        }

        public static void WorkbookToCsv(string pathToExcel, string pathToWrite)
        {
            Excel.Application excel = new Excel.Application();
            excel.DisplayAlerts = false;
            var workbook = excel.Workbooks.Open(pathToExcel, false, true, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, false, true, Type.Missing);
            string workbookName = workbook.Name;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in workbook.Sheets)
            {
                sheet.SaveAs(pathToWrite + "\\" + workbookName + "." + sheet.Name + ".csv", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
            }

            workbook.Close(false);
            excel.Quit();
            workbook = null;
            excel = null;
        }
    }
}
