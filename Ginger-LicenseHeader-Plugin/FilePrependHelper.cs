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
            try
            {
                string strFileContents = ReadToEndSkipInitialComment(initialComment);
                strFileContents = initialComment + Environment.NewLine + strDataToAppend + Environment.NewLine + strFileContents;
                using (StreamWriter swWriter = new StreamWriter(file, false))
                {
                    swWriter.Write(strFileContents);
                    swWriter.Close();
                }
            }
            catch (Exception objException)
            {
                throw (objException);
            }
        }

        public void Prepend(string strDataToAppend)
        {
            try
            {
                string strFileContents = ReadFileToEnd();
                strFileContents = strDataToAppend + Environment.NewLine + strFileContents;
                using (StreamWriter swWriter = new StreamWriter(file, false))
                {
                    swWriter.Write(strFileContents);
                    swWriter.Close();
                }
            }
            catch (Exception objException)
            {
                throw (objException);
            }
        }

        private string ReadFileToEnd()
        {
            string strFileContents;
            using (StreamReader sr = File.OpenText(file))
            {
                strFileContents = sr.ReadToEnd();
                sr.Close();
            }
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

        public void ReplaceText(string findText, string replaceText)
        {
            try
            {
                string strFileContents = GetReplacedText(findText, replaceText);
                if (!string.IsNullOrEmpty(strFileContents))
                {
                    using (StreamWriter swWriter = new StreamWriter(file, false))
                    {
                        swWriter.Write(strFileContents);
                        swWriter.Close();
                    }
                }
            }
            catch (Exception objException)
            {
                throw (objException);
            }
        }

        private string GetReplacedText(string findText, string replaceText)
        {
            string strFileContents = string.Empty;
            using (StreamReader sr = File.OpenText(file))
            {
                strFileContents = sr.ReadToEnd();
            }

            if (!string.IsNullOrEmpty(strFileContents) && strFileContents.Contains(findText))
            {
                strFileContents = strFileContents.Replace(findText, replaceText);
            }
            else
            {
                strFileContents = string.Empty;
            }
            return strFileContents;
        }
    }
}
