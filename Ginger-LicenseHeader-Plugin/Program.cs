using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Ginger_LicenseHeader_Plugin
{
    class Program
    {
        private const string FILE_REVIEW_COMMENT_STATUS = "//# Status=";
        private const string XAML_FILE_REVIEW_COMMENT_STATUS = "<!--//# Status=";
        private const string FILE_CONSTANT_COMMENT = "Comment=";

        const string HEADER_LICENSE_LINE = "#region License";
        private const string ZAMMEL_FILE = "Zammel File";
        private const string JAVA_FILE = "Java File";
        private const string PDF_FILE = "PDF File";
        private const string PNG_FILE = "Png File";
        private const string CODE_FILE = "Code File";
        private const string JS_FILE = "Java Script File";
        private const string REVIEW_COMMENT_STATUS = "Reviewed";
        private const string CLEANED_COMMENT_STATUS = "Cleaned";

        static List<string> projects = new List<string>
        {
            @"\Ginger\", @"\GingerAssemblyUpdater\", @"\GingerATS\", @"\GingerConsole\", @"\GingerCore\", @"\GingerCoreNET\",
            @"\GingerHelper\", @"\GingerPACTPlugIn\", @"\GingerPlugIns\", @"\GingerPlugInsNET\", @"\GingerWebServicesPlugin\", @"\GingerWebServicesPluginWPF\",
            @"\GingerWPF\", @"\GingerWPFDriverWindow\", @"\SeleniumPlugin\", @"\SeleniumPluginWPF\", @"\StandAloneActions\", @"\UIAComWrapper\",
            @"\GingerRemoteAgent\", @"\GingerCodeGen\", @"\GingerCorePlugin\", @"\GingerFunctionsRepository\", @"\GingerRemoteAgent\",
            @"\GingerTools\", @"\SeleniumDriverCP\", @"\GingerQC\"
        };

        static List<string> utprojects = new List<string>
        {
            @"\GingerConsoleTest\", @"\GingerCoreNETUnitTest\",
            @"\GingerWebServicesPluginTest\", @"\GingerWPFDriverWindowTest\", @"\GingerWPFUnitTest\",
            @"\SeleniumPluginTest\", @"\StandAloneActionsTest\", @"\UnitTests\", @"\UnitTestsCP\",
            @"\UIAutomationTests\",@"\SeleniumDriverCPUnitTests\", @"\GingerUnitTester\",@"\GingerConsoleUnitTest\",

        };

        static void Main(string[] args)
        {
            bool addComments = true;
            if (addComments)
            {
                AddCommentsToGingerFiles();
            }
            else
            {
                GetGingerFiles();
                //GetGingerAllFilesList();
            }
        }

        private static void AddCommentsToGingerFiles()
        {
            try
            {
                string HEADER_COMMENT = "#region License" + Environment.NewLine +
                                              "/*" + Environment.NewLine +
                                              "Copyright © 2014-2018 European Support Limited" + Environment.NewLine + Environment.NewLine +
                                              "Licensed under the Apache License, Version 2.0 (the \"License\")" + Environment.NewLine +
                                              "you may not use this file except in compliance with the License." + Environment.NewLine +
                                              "You may obtain a copy of the License at " + Environment.NewLine + Environment.NewLine +
                                              "http://www.apache.org/licenses/LICENSE-2.0 " + Environment.NewLine + Environment.NewLine +
                                              "Unless required by applicable law or agreed to in writing, software" + Environment.NewLine +
                                              "distributed under the License is distributed on an \"AS IS\" BASIS, " + Environment.NewLine +
                                              "WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. " + Environment.NewLine +
                                              "See the License for the specific language governing permissions and " + Environment.NewLine +
                                              "limitations under the License. " + Environment.NewLine +
                                              "*/" + Environment.NewLine +
                                              "#endregion" + Environment.NewLine;

                //string path = @"C:\Ginger\GingerDirectoryScanner\DemoFileOperation\DemoFileOperation";
                //string path = @"C:\Ginger\GingerNextVer_Dev\";
                string path = @"C:\GitHubRepository\Ginger\Ginger";
                var ext = new List<string> { ".cs" };
                var filePaths = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories)
                    .Where(file => ext
                        .Contains(Path.GetExtension(file)))
                    .ToList();

                LogText(MethodBase.GetCurrentMethod().Name, string.Format("===============================Started Adding License Comments, Project - {0}===============================", path));

                foreach (var filepath in filePaths)
                {
                    if (!filepath.Contains("Temporary")
                        && !filepath.Contains(".g.") && !filepath.Contains(".g.i.")
                        && !filepath.Contains("AssemblyInfo") && !filepath.ToLower().Contains(".designer")
                        && !filepath.ToLower().Contains("\\bin") && !filepath.ToLower().Contains("\\obj"))
                    {
                        bool isUtProject = IsUTProject(filepath.Substring(path.Length));
                        if (!isUtProject)
                        {
                            string initialComment = string.Empty;
                            bool isLicenseCommentPresent = IsLicenseHeaderPresent(filepath, ref initialComment);
                            if (!isLicenseCommentPresent)
                            {
                                FilePrependHelper fp = new FilePrependHelper(filepath);
                                if (string.IsNullOrEmpty(initialComment))
                                {
                                    fp.Prepend(HEADER_COMMENT);
                                }
                                else
                                {
                                    fp.PrependTextKeepInitialComment(HEADER_COMMENT, initialComment);
                                }
                            }

                            LogText(MethodBase.GetCurrentMethod().Name, string.Format("Earlier - {0}, License Added - {1}, FileName - {2}", isLicenseCommentPresent, !isLicenseCommentPresent, filepath));
                        }
                    }
                }

                LogText(MethodBase.GetCurrentMethod().Name, string.Format("===============================Ended Adding License Comments, Project - {0}===============================", path));
                LogText(MethodBase.GetCurrentMethod().Name, string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.StackTrace);
            }
        }

        private static bool IsLicenseHeaderPresent(string filepath, ref string initialComment)
        {
            bool isLicenseCommentPresent = false;
            try
            {
                using (StreamReader sr = File.OpenText(filepath))
                {
                    int count = 1;
                    string str = String.Empty;


                    while ((str = sr.ReadLine()) != null)
                    {
                        if (HEADER_LICENSE_LINE == str)
                        {
                            isLicenseCommentPresent = true;
                            break;
                        }

                        if (str.StartsWith(FILE_REVIEW_COMMENT_STATUS))
                        {
                            initialComment = str;
                        }

                        if (count >= 2)
                        {
                            break;
                        }
                        count++;
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return isLicenseCommentPresent;
        }

        private static void GetGingerFilesTemp()
        {
            try
            {
                string path = @"C:\Ginger\GingerNextVer_Dev";
                var ext = new List<string> { ".cs", ".xaml", ".java", ".pdf", ".js" };     //, ".xaml"
                var filePaths = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories)
                    .Where(file => ext
                        .Contains(Path.GetExtension(file)))
                    .ToList();

                Application oXL;
                _Workbook oWB;
                _Worksheet oSheet;
                _Worksheet oSheet2;
                Range oRng;
                Range oRng2;
                int rowIndex = 2;

                //Start Excel and get Application object.
                oXL = new Application();

                double totalFilesCount = 0;
                double cleanedTotalFilesCount = 0;
                double reviewedTotalFilesCount = 0;

                double cleanedCodeFilesCount = 0;
                double reviewedCodeFilesCount = 0;
                double totalCodeFilesCount = 0;

                double totalPdfFilesCount = 0;
                double totalPngFilesCount = 0;

                double cleanedZammelFilesCount = 0;
                double reviewedZammelFilesCount = 0;
                double totalZammelFilesCount = 0;

                double cleanedJavaFilesCount = 0;
                double reviewedJavaFilesCount = 0;
                double totalJavaFilesCount = 0;

                double cleanedJavaScriptFilesCount = 0;
                double reviewedJavaScriptFilesCount = 0;
                double totalJavaScriptFilesCount = 0;

                //Get a new workbook.
                oWB = oXL.Workbooks.Add("");
                oSheet = (_Worksheet)oWB.ActiveSheet;
                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "File Name";
                oSheet.Cells[1, 2] = "Project";
                oSheet.Cells[1, 3] = "File Type";
                oSheet.Cells[1, 4] = "Timestamp";
                oSheet.Cells[1, 5] = "Status";
                oSheet.Cells[1, 6] = "Comments";
                oSheet.Cells[1, 7] = "File Path";
                oSheet.Cells[1, 8] = "License Header Present";
                oSheet.Cells[1, 9] = "File Size";
                foreach (var filepath in filePaths)
                {
                    if (!filepath.Contains("Temporary")
                        && !filepath.Contains(".g.") && !filepath.Contains(".g.i.")
                        && !filepath.Contains("AssemblyInfo") && !filepath.ToLower().Contains(".designer"))
                    {
                        bool isUtProject = IsUTProject(filepath.Substring(path.Length));
                        string project = GetProject(filepath.Substring(path.Length));

                        string initialComment = string.Empty;
                        bool isLicenseCommentPresent = IsLicenseHeaderPresent(filepath, ref initialComment);

                        if (!string.IsNullOrEmpty(project) && !isUtProject)
                        {
                            FileInfo fi = new FileInfo(filepath);
                            string filename = fi.Name;
                            string filetype = GetFileType(fi.Extension);
                            DateTime timestamp = fi.CreationTime;
                            string line = File.ReadLines(filepath).First();
                            string status = string.Empty;
                            string comment = string.Empty;

                            if (!string.IsNullOrEmpty(line) && line.StartsWith(FILE_REVIEW_COMMENT_STATUS))
                            {
                                string[] str = line.Split(';');
                                if (str != null && str.Length > 0)
                                {
                                    status = str[0].Substring(FILE_REVIEW_COMMENT_STATUS.Length);
                                    comment = str[1].Replace("=", "").Substring(FILE_CONSTANT_COMMENT.Length);
                                }
                            }
                            ////--------Text File write-----------
                            //using (System.IO.StreamWriter file = new System.IO.StreamWriter(Path.Combine(path, "FilesList.txt"), true))
                            //{
                            //    file.WriteLine($"Project: {project}, File: {filename}, Type: {filetype}, Path: {filepath}, Last Modified: {timestamp}, Status: {status}, Comment: {comments}");
                            //}
                            ////--------Text File write-----------

                            //--------Excel File write-----------
                            try
                            {
                                oSheet.get_Range("A1", "G1").Font.Bold = true;
                                oSheet.get_Range("A1", "G1").VerticalAlignment = XlVAlign.xlVAlignCenter;

                                oSheet.Cells[rowIndex, 1] = filename;
                                oSheet.Cells[rowIndex, 2] = project;
                                oSheet.Cells[rowIndex, 3] = filetype;
                                oSheet.Cells[rowIndex, 4] = Convert.ToString(timestamp.ToString("dd-MM-yyyy hh:mm:ss"));
                                oSheet.Cells[rowIndex, 5] = status;
                                oSheet.Cells[rowIndex, 6] = comment;
                                oSheet.Cells[rowIndex, 7] = filepath;
                                oSheet.Cells[rowIndex, 8] = isLicenseCommentPresent;
                                oSheet.Cells[rowIndex, 9] = fi.Length;

                                //AutoFit columns A:G.
                                oRng = oSheet.get_Range("A1", "I1");
                                oRng.EntireColumn.AutoFit();

                                rowIndex++;
                                //if (rowIndex == 10)
                                //{
                                //    break;
                                //}
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.StackTrace);
                            }

                            //--------Excel File write----------- 

                            totalFilesCount += 1;
                            if (filetype == ZAMMEL_FILE)
                            {
                                if (status == REVIEW_COMMENT_STATUS)
                                {
                                    reviewedZammelFilesCount += 1;
                                }
                                else if (status == CLEANED_COMMENT_STATUS)
                                {
                                    cleanedZammelFilesCount += 1;
                                }
                                totalZammelFilesCount += 1;
                            }
                            else if (filetype == CODE_FILE)
                            {
                                if (status == REVIEW_COMMENT_STATUS)
                                {
                                    reviewedCodeFilesCount += 1;
                                }
                                else if (status == CLEANED_COMMENT_STATUS)
                                {
                                    cleanedCodeFilesCount += 1;
                                }
                                totalCodeFilesCount += 1;
                            }
                            else if (filetype == JS_FILE)
                            {
                                if (status == REVIEW_COMMENT_STATUS)
                                {
                                    reviewedJavaScriptFilesCount += 1;
                                }
                                else if (status == CLEANED_COMMENT_STATUS)
                                {
                                    cleanedJavaScriptFilesCount += 1;
                                }
                                totalJavaScriptFilesCount += 1;
                            }
                            else if (filetype == JAVA_FILE)
                            {
                                if (status == REVIEW_COMMENT_STATUS)
                                {
                                    reviewedJavaFilesCount += 1;
                                }
                                else if (status == CLEANED_COMMENT_STATUS)
                                {
                                    cleanedJavaFilesCount += 1;
                                }
                                totalJavaFilesCount += 1;
                            }
                            else if (filetype == PDF_FILE)
                            {
                                totalPdfFilesCount += 1;
                            }
                            else if (filetype == PNG_FILE)
                            {
                                totalPngFilesCount += 1;
                            }
                        }
                    }
                }

                double percent = 0;
                oWB.Worksheets.Add();
                oSheet2 = (_Worksheet)oWB.ActiveSheet;

                oSheet2.get_Range("A1", "F1").Font.Bold = true;
                oSheet2.get_Range("A1", "F1").VerticalAlignment = XlVAlign.xlVAlignCenter;

                oSheet2.Cells[1, 1] = "Files";
                oSheet2.Cells[1, 2] = "Total Files";
                oSheet2.Cells[1, 3] = "Cleaned Files";
                oSheet2.Cells[1, 4] = "Cleaned Percentage";
                oSheet2.Cells[1, 5] = "Reviewed Files";
                oSheet2.Cells[1, 6] = "Reviewed Percentage";

                oSheet2.Cells[2, 1] = CODE_FILE;
                oSheet2.Cells[2, 2] = totalCodeFilesCount;
                oSheet2.Cells[2, 3] = cleanedCodeFilesCount;
                percent = (cleanedCodeFilesCount / totalCodeFilesCount) * 100;
                oSheet2.Cells[2, 4] = percent;
                oSheet2.Cells[2, 5] = reviewedCodeFilesCount;
                percent = (reviewedCodeFilesCount / totalCodeFilesCount) * 100;
                oSheet2.Cells[2, 6] = percent;

                oSheet2.Cells[3, 1] = ZAMMEL_FILE;
                oSheet2.Cells[3, 2] = totalZammelFilesCount;
                oSheet2.Cells[3, 3] = cleanedZammelFilesCount;
                percent = (cleanedZammelFilesCount / totalZammelFilesCount) * 100;
                oSheet2.Cells[3, 4] = percent;
                oSheet2.Cells[3, 5] = reviewedZammelFilesCount;
                percent = (reviewedZammelFilesCount / totalZammelFilesCount) * 100;
                oSheet2.Cells[3, 6] = percent;

                oSheet2.Cells[4, 1] = JS_FILE;
                oSheet2.Cells[4, 2] = totalJavaScriptFilesCount;
                oSheet2.Cells[4, 3] = cleanedJavaScriptFilesCount;
                percent = (cleanedJavaScriptFilesCount / totalJavaScriptFilesCount) * 100;
                oSheet2.Cells[4, 4] = percent;
                oSheet2.Cells[4, 5] = reviewedJavaScriptFilesCount;
                percent = (reviewedJavaScriptFilesCount / totalJavaScriptFilesCount) * 100;
                oSheet2.Cells[4, 6] = percent;

                oSheet2.Cells[5, 1] = JAVA_FILE;
                oSheet2.Cells[5, 2] = totalJavaFilesCount;
                oSheet2.Cells[5, 3] = cleanedJavaFilesCount;
                percent = (cleanedJavaFilesCount / totalJavaFilesCount) * 100;
                oSheet2.Cells[5, 4] = percent;
                oSheet2.Cells[5, 5] = reviewedJavaFilesCount;
                percent = (reviewedJavaFilesCount / totalJavaFilesCount) * 100;
                oSheet2.Cells[5, 6] = percent;

                oSheet2.Cells[6, 1] = PDF_FILE;
                oSheet2.Cells[6, 2] = totalPdfFilesCount;
                oSheet2.Cells[6, 3] = 0;
                oSheet2.Cells[6, 4] = 0;
                oSheet2.Cells[6, 5] = 0;
                oSheet2.Cells[6, 6] = 0;

                oSheet2.Cells[7, 1] = PNG_FILE;
                oSheet2.Cells[7, 2] = totalPdfFilesCount;
                oSheet2.Cells[7, 3] = 0;
                oSheet2.Cells[7, 4] = 0;
                oSheet2.Cells[7, 5] = 0;
                oSheet2.Cells[7, 6] = 0;

                oSheet2.Cells[8, 1] = "All Files";
                oSheet2.Cells[8, 2] = totalFilesCount;
                oSheet2.Cells[8, 3] = (cleanedCodeFilesCount + cleanedZammelFilesCount + cleanedJavaFilesCount + cleanedJavaScriptFilesCount);
                percent = ((cleanedCodeFilesCount + cleanedZammelFilesCount + cleanedJavaFilesCount + cleanedJavaScriptFilesCount) / totalFilesCount) * 100;
                oSheet2.Cells[8, 4] = percent;
                oSheet2.Cells[8, 5] = (reviewedCodeFilesCount + reviewedZammelFilesCount + reviewedJavaFilesCount + reviewedJavaScriptFilesCount);
                percent = ((reviewedCodeFilesCount + reviewedZammelFilesCount + reviewedJavaFilesCount + reviewedJavaScriptFilesCount) / totalFilesCount) * 100;
                oSheet2.Cells[8, 6] = percent;

                //AutoFit columns A:G.
                oRng2 = oSheet2.get_Range("A1", "F1");
                oRng2.EntireColumn.AutoFit();

                oXL.Visible = false;
                oXL.UserControl = false;

                oWB.SaveAs(Path.Combine(path, $"FilesList_{DateTime.Now:dd-MM-yyyy}.xlsx"),
                    XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                oWB.Close();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.StackTrace);
            }
        }

        private static void GetGingerFiles()
        {
            int rowIndex = 2;
            try
            {
                string prjPath = @"C:\Ginger\GingerNextVer_Dev";
                string path = @"C:\Ginger\GingerNextVer_Dev";
                //var filePaths = Directory.GetFiles(prjPath, "*.*", SearchOption.AllDirectories).ToList();
                var ext = new List<string>
                {
                    ".baml", ".cache", ".dylib", ".lref", ".pdb", ".resources", ".so", ".svnExe", ".ide",
                    ".ide", ".suo", ".csproj", ".vspscc", ".settings", ".properties", ".vssscc", ".sln",
                    ".Cache", ".resx", ".shfbproj", ".user", ""
                };
                var filePaths = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories)
                    .Where(file => !ext.Contains(Path.GetExtension(file)))
                    .ToList();

                Application oXL;
                _Workbook oWB;
                _Worksheet oSheet;
                _Worksheet oSheet2;
                Range oRng;
                Range oRng2;

                //Start Excel and get Application object.
                oXL = new Application();

                //Get a new workbook.
                oWB = oXL.Workbooks.Add("");
                oSheet = (_Worksheet)oWB.ActiveSheet;
                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "File Name";
                oSheet.Cells[1, 2] = "Project";
                oSheet.Cells[1, 3] = "File Type";
                oSheet.Cells[1, 4] = "Timestamp";
                oSheet.Cells[1, 5] = "Status";
                oSheet.Cells[1, 6] = "Comments";
                oSheet.Cells[1, 7] = "File Path";
                oSheet.Cells[1, 8] = "License Header Present";
                oSheet.Cells[1, 9] = "File Size";

                List<FileDetails> fileDetList = new List<FileDetails>();
                foreach (var filepath in filePaths)
                {
                    if (!filepath.Contains("Temporary")
                        && !filepath.Contains(".g.") && !filepath.Contains(".g.i.")
                        && !filepath.Contains("AssemblyInfo") && !filepath.ToLower().Contains(".designer")
                        && !filepath.ToLower().Contains("\\bin") && !filepath.ToLower().Contains("\\obj"))
                    {
                        bool isUtProject = IsUTProject(filepath.Substring(prjPath.Length));
                        string project = GetProject(filepath.Substring(path.Length));

                        string initialComment = string.Empty;

                        if (!string.IsNullOrEmpty(project) && !isUtProject)
                        {
                            bool isReadOnlyFile = false;
                            string line = String.Empty;
                            bool isLicenseCommentPresent = false;
                            try
                            {
                                line = File.ReadLines(filepath).First();
                                isLicenseCommentPresent = IsLicenseHeaderPresent(filepath, ref initialComment);
                            }
                            catch (Exception ex)
                            {
                                isReadOnlyFile = true;
                            }

                            if (!isReadOnlyFile)
                            {
                                FileInfo fi = new FileInfo(filepath);

                                string filename = fi.Name;
                                string filetype = fi.Extension;
                                DateTime timestamp = fi.CreationTime;

                                string status = string.Empty;
                                string comment = string.Empty;

                                if (!string.IsNullOrEmpty(line) && line.StartsWith(FILE_REVIEW_COMMENT_STATUS))
                                {
                                    string[] str = line.Split(';');
                                    if (str != null && str.Length > 0)
                                    {
                                        status = str[0].Substring(FILE_REVIEW_COMMENT_STATUS.Length);
                                        comment = str[1].Replace("=", "").Substring(FILE_CONSTANT_COMMENT.Length);
                                    }
                                }
                                else if (!string.IsNullOrEmpty(line) && line.StartsWith(XAML_FILE_REVIEW_COMMENT_STATUS))
                                {
                                    string[] str = line.Split(';');
                                    if (str != null && str.Length > 0)
                                    {
                                        status = str[0].Substring(XAML_FILE_REVIEW_COMMENT_STATUS.Length);
                                        comment = str[1].Replace("=", "").Replace("-->", "").Substring(FILE_CONSTANT_COMMENT.Length);
                                    }
                                }

                                if (filetype == ".mht" || filetype == ".gradle" || filetype == ".iml"
                                    || filetype == ".pro" || filetype == ".jelly" || filetype == ".il"
                                    || filetype == ".iss" || filetype == ".gitignore")
                                {
                                    status = "Keep";
                                    comment = "Yaron to review if can be deleted";
                                }
                                ////--------Text File write-----------
                                //using (System.IO.StreamWriter file = new System.IO.StreamWriter(Path.Combine(path, "FilesList.txt"), true))
                                //{
                                //    file.WriteLine($"Project: {project}, File: {filename}, Type: {filetype}, Path: {filepath}, Last Modified: {timestamp}, Status: {status}, Comment: {comments}");
                                //}
                                ////--------Text File write-----------

                                //--------Excel File write-----------
                                try
                                {
                                    oSheet.get_Range("A1", "G1").Font.Bold = true;
                                    oSheet.get_Range("A1", "G1").VerticalAlignment = XlVAlign.xlVAlignCenter;

                                    oSheet.Cells[rowIndex, 1] = filename;
                                    oSheet.Cells[rowIndex, 2] = project;
                                    oSheet.Cells[rowIndex, 3] = filetype;
                                    oSheet.Cells[rowIndex, 4] = Convert.ToString(timestamp.ToString("dd-MM-yyyy hh:mm:ss"));
                                    oSheet.Cells[rowIndex, 5] = status;
                                    oSheet.Cells[rowIndex, 6] = comment;
                                    oSheet.Cells[rowIndex, 7] = filepath;
                                    oSheet.Cells[rowIndex, 8] = isLicenseCommentPresent;
                                    oSheet.Cells[rowIndex, 9] = fi.Length;

                                    //AutoFit columns A:G.
                                    oRng = oSheet.get_Range("A1", "I1");
                                    oRng.EntireColumn.AutoFit();

                                    rowIndex++;
                                    //if (rowIndex == 20)
                                    //{
                                    //    break;
                                    //}

                                    if (filetype == ".mht" || filetype == ".gradle" || filetype == ".iml"
                                    || filetype == ".pro" || filetype == ".jelly" || filetype == ".il"
                                    || filetype == ".iss" || filetype == ".gitignore")
                                    {
                                        status = string.Empty;
                                        comment = string.Empty;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine(ex.StackTrace);
                                }

                                //--------Excel File write----------- 

                                if (!IsFileTypeExists(fileDetList, filetype))
                                {
                                    fileDetList.Add(new FileDetails
                                    {
                                        FileType = filetype
                                    });
                                }

                                UpdateFileCount(fileDetList, filetype, status);
                            }
                        }
                    }
                }

                oWB.Worksheets.Add();
                oSheet2 = (_Worksheet)oWB.ActiveSheet;

                oSheet2.get_Range("A1", "F1").Font.Bold = true;
                oSheet2.get_Range("A1", "F1").VerticalAlignment = XlVAlign.xlVAlignCenter;

                oSheet2.Cells[1, 1] = "Files";
                oSheet2.Cells[1, 2] = "Total Files";
                oSheet2.Cells[1, 3] = "Cleaned Files";
                oSheet2.Cells[1, 4] = "Cleaned Percentage";
                oSheet2.Cells[1, 5] = "Reviewed Files";
                oSheet2.Cells[1, 6] = "Reviewed Percentage";

                int count = 2;
                UpdateFilePercentage(fileDetList);
                foreach (var item in fileDetList)
                {
                    oSheet2.Cells[count, 1] = item.FileType;
                    oSheet2.Cells[count, 2] = item.TotalFilesCount;
                    oSheet2.Cells[count, 3] = item.CleanedCount;
                    oSheet2.Cells[count, 4] = item.CleanedPercent;
                    oSheet2.Cells[count, 5] = item.ReviewedCount;
                    oSheet2.Cells[count, 6] = item.ReviewedPercent;
                    count += 1;
                }

                //AutoFit columns A:G.
                oRng2 = oSheet2.get_Range("A1", "F1");
                oRng2.EntireColumn.AutoFit();

                oXL.Visible = false;
                oXL.UserControl = false;

                oWB.SaveAs(Path.Combine(path, $"AllFilesList_{DateTime.Now:dd-MM-yyyy}.xlsx"),
                    XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                oWB.Close();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.StackTrace);
                Console.WriteLine(Convert.ToString(rowIndex) + ex.StackTrace);
            }
        }

        private static void UpdateFilePercentage(List<FileDetails> fileDetList)
        {
            if (fileDetList != null && fileDetList.Count > 0)
            {
                double totalFilesCount = 0;
                double reviewedCount = 0;
                double cleanedCount = 0;
                foreach (FileDetails item in fileDetList)
                {
                    item.ReviewedPercent = (item.ReviewedCount / item.TotalFilesCount) * 100;
                    item.CleanedPercent = (item.CleanedCount / item.TotalFilesCount) * 100;
                    totalFilesCount += item.TotalFilesCount;
                    reviewedCount += item.ReviewedCount;
                    cleanedCount += item.CleanedCount;
                }

                fileDetList.Add(new FileDetails
                {
                    FileType = "All Files",
                    TotalFilesCount = totalFilesCount,
                    ReviewedCount = reviewedCount,
                    CleanedCount = cleanedCount,
                    ReviewedPercent = ((reviewedCount / totalFilesCount) * 100),
                    CleanedPercent = ((cleanedCount / totalFilesCount) * 100)
                });
            }
        }

        private static void UpdateFileCount(List<FileDetails> fileDetList, string filetype, string status)
        {
            if (fileDetList != null && fileDetList.Count > 0)
            {
                foreach (FileDetails item in fileDetList)
                {
                    if (item.FileType == filetype)
                    {
                        item.TotalFilesCount += 1;
                        if (status == REVIEW_COMMENT_STATUS)
                        {
                            item.ReviewedCount += 1;
                        }
                        else if (status == CLEANED_COMMENT_STATUS)
                        {
                            item.CleanedCount += 1;
                        }
                    }

                    if (filetype == ".pdf" || filetype == ".dgml" || filetype == ".docx" || filetype == ".ttf"
                        || filetype == ".zip" || filetype == ".xml" || filetype == ".config" || filetype == ".ico"
                        || filetype == ".css" || filetype == ".xshd" || filetype == ".txt" || filetype == ".gradle"
                        || filetype == ".robot" || filetype == ".py" || filetype == ".json" || filetype == ".bat"
                        || filetype == ".vbs" || filetype == ".rtf" || filetype == ".exe" || filetype == ".jar"
                        || filetype == ".eot" || filetype == ".jpg" || filetype == ".gif" || filetype == ".bmp"
                        || filetype == ".manifest" || filetype == ".pro" || filetype == ".apk" || filetype == ".rawproto"
                        || filetype == ".classpath" || filetype == ".project" || filetype == ".prefs" || filetype == ".jardesc"
                        || filetype == ".MF" || filetype == ".il" || filetype == ".mht" || filetype == ".mdb"
                        || filetype == ".mdf" || filetype == ".ldf" || filetype == ".jelly" || filetype == ".gitignore"
                        || filetype == ".iss" || filetype == ".iml")
                    {
                        if (item.FileType == filetype)
                        {
                            item.CleanedCount += 1;
                        }
                    }
                }
            }
        }

        private static bool IsFileTypeExists(List<FileDetails> fileDetList, string filetype)
        {
            bool isExists = false;
            if (fileDetList != null && fileDetList.Count > 0)
            {
                foreach (FileDetails details in fileDetList)
                {
                    if (details.FileType.ToLower() == filetype.ToLower())
                    {
                        isExists = true;
                        break;
                    }
                }
            }
            return isExists;
        }

        private static void GetGingerAllFilesList()
        {
            try
            {
                string path = @"C:\Ginger\GingerNextVer_Dev";
                var filePaths = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories)
                    .ToList();

                Application oXL;
                _Workbook oWB;
                _Worksheet oSheet;
                int rowIndex = 2;

                //Start Excel and get Application object.
                oXL = new Application();

                //Get a new workbook.
                oWB = oXL.Workbooks.Add("");
                oSheet = (_Worksheet)oWB.ActiveSheet;
                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "File Name";
                foreach (var filepath in filePaths)
                {
                    oSheet.Cells[rowIndex, 1] = filepath;
                    rowIndex += 1;
                }

                oXL.Visible = false;
                oXL.UserControl = false;

                oWB.SaveAs(Path.Combine(path, $"AllFilesList_{DateTime.Now:dd-MM-yyyy}.xlsx"),
                    XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                oWB.Close();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.StackTrace);
            }
        }

        private static bool IsUTProject(string filepath)
        {
            bool isUTProject = false;
            foreach (string path in utprojects)
            {
                if (filepath.Contains(path))
                {
                    isUTProject = true;
                    break;
                }
            }
            return isUTProject;
        }

        private static string GetProject(string filepath)
        {
            string project = string.Empty;
            foreach (string path in projects)
            {
                if (filepath.Contains(path))
                {
                    project = path.Replace(@"\", "");
                    break;
                }
            }

            if (project == "GingerTools" || project == "GingerCodeGen" || project == "GingerFunctionsRepository")
            {
                project = string.Empty;
            }
            return project;
        }

        private static string GetFileType(string fiExtension)
        {
            string fiType = string.Empty;
            switch (fiExtension)
            {
                case ".cs":
                    fiType = CODE_FILE;
                    break;
                case ".xaml":
                    fiType = ZAMMEL_FILE;
                    break;
                case ".java":
                    fiType = JAVA_FILE;
                    break;
                case ".js":
                    fiType = JS_FILE;
                    break;
                case ".pdf":
                    fiType = PDF_FILE;
                    break;
                case ".png":
                    fiType = PNG_FILE;
                    break;
            }
            return fiType;
        }

        private static void LogText(string methodName, string text)
        {
            string file = methodName + "_log.txt";
            using (StreamWriter swWriter = File.AppendText(file))
            {
                if (!string.IsNullOrEmpty(text))
                {
                    swWriter.WriteLine(string.Format("{0} - {1}", DateTime.Now.ToString("dd-MM-yyy hh:mm:ss"), text));
                }
                else
                {
                    swWriter.WriteLine(string.Format("{0}", text));
                }
                swWriter.Flush();
            }
        }
    }
}
