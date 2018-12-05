using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Ginger_LicenseHeader_Plugin
{
    public class FilePrependHelper
    {
        private string file = null;
        public FilePrependHelper(string filePath)
        {
            file = filePath;
        }

        public void PrependTextKeepInitialComment(string strDataToAppend, string initialComment)
        {
            string strFileContents = "";
            StreamWriter swWriter = null;
            try
            {
                strFileContents = ReadToEndSkipInitialComment(initialComment);

                strFileContents = initialComment + Environment.NewLine + strDataToAppend + Environment.NewLine + strFileContents;
                swWriter = new StreamWriter(file, false);
                swWriter.Write(strFileContents);
                swWriter.Flush();
            }
            catch (Exception objException)
            {
                throw (objException);
            }
            finally
            {
                swWriter.Close();
            }
        }

        public void Prepend(string strDataToAppend)
        {
            string strFileContents = "";
            StreamWriter swWriter = null;
            try
            {
                strFileContents = ReadFileToEnd();

                strFileContents = strDataToAppend + Environment.NewLine + strFileContents;
                swWriter = new StreamWriter(file, false);
                swWriter.Write(strFileContents);
                swWriter.Flush();
            }
            catch (Exception objException)
            {
                throw (objException);
            }
            finally
            {
                swWriter.Close();
            }
        }

        private string ReadFileToEnd()
        {
            string strFileContents;
            StreamReader srReader = new StreamReader(file);
            strFileContents = srReader.ReadToEnd();
            srReader.Close();
            return strFileContents;
        }

        private string ReadToEndSkipInitialComment(string initialComment)
        {
            StringBuilder strFileContents = new StringBuilder();
            using (StreamReader sr = File.OpenText(file))
            {
                string str = String.Empty;
                while ((str = sr.ReadLine()) != null)
                {
                    if (str != initialComment)
                    {
                        strFileContents.Append(str + Environment.NewLine);
                    }
                }
            }
            return strFileContents.ToString();
        }
    }
}
