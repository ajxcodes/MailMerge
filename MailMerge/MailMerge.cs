using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Xml.Linq;

namespace MailMerge
{
    /// <summary>
    /// Base class for mailmerging Word documents programmatically.
    /// </summary>
    public class MailMerge
    {
        private string _templateFilePath;
        private string _targetFilePath;
        private byte[] _targetFileBytes;
        private List<byte[]> _targetFilesBytes;
        private const string FieldDelimeter = " MERGEFIELD ";

        /// <summary>
        /// Outputs a true success if mail merge completes successfully. Otherwise, an exception will be stored in the Exception field
        /// </summary>
        public struct RESULT
        {
            public bool Successful;
            public string Exception;
        }

        public MailMerge(string templatePath, string targetPath)
        {
            _templateFilePath = templatePath;
            _targetFilePath = targetPath;
        }

        /// <summary>
        /// Converts a .DOTX template to a .DOCX Word document
        /// </summary>
        /// <returns></returns>
        public RESULT ConvertTemplate()
        {
            try
            {
                MemoryStream msFile = null;

                using (Stream sTemplate = File.Open(_templateFilePath, FileMode.Open, FileAccess.Read))
                {
                    msFile = new MemoryStream((int)sTemplate.Length);
                    sTemplate.CopyTo(msFile);
                    msFile.Position = 0L;
                }

                using (WordprocessingDocument wpdFile = WordprocessingDocument.Open(msFile, true))
                {
                    wpdFile.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

                    MainDocumentPart docPart = wpdFile.MainDocumentPart;
                    docPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate", new Uri(_templateFilePath, UriKind.RelativeOrAbsolute));

                    docPart.Document.Save();
                }

                File.WriteAllBytes(_targetFilePath, msFile.ToArray());

                msFile.Close();

                return new RESULT { Successful = true };
            }
            catch(Exception e)
            {
                return new RESULT { Successful = false, Exception = e.ToString() };
            }
        }

        /// <summary>
        /// Retrieves the merge values based on the field name. This method will be overriden in child class.
        /// </summary>
        /// <param name="FieldName"></param>
        /// <returns>Merge Value</returns>
        public virtual string GetMergeValue(string FieldName)
        {
            switch (FieldName)
            {
                default:
                    throw new Exception(message: FieldName + " field not found.");
            }
        }
        /// <summary>
        /// Processes the mail merge and saves the bytes. This method will be overriden in child class.
        /// </summary>
        /// <returns></returns>
        public virtual RESULT Process()
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(_targetFilePath, true))
                {
                    foreach (FieldCode field in doc.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {
                        var fieldNameStart = field.Text.LastIndexOf(FieldDelimeter, System.StringComparison.Ordinal);
                        var fieldname = field.Text.Substring(fieldNameStart + FieldDelimeter.Length).Trim();
                        var fieldValue = GetMergeValue(FieldName: fieldname);

                        foreach (Run run in doc.MainDocumentPart.Document.Descendants<Run>())
                        {
                            foreach (Text txtFromRun in run.Descendants<Text>().Where(a => a.Text == "«" + fieldname + "»"))
                            {
                                txtFromRun.Text = fieldValue;
                            }
                        }
                    }
                    _targetFileBytes = File.ReadAllBytes(_targetFilePath);
                    _targetFilesBytes.Add(_targetFileBytes);
                    return new RESULT { Successful = true };
                }
            }
            catch(Exception e)
            {
                return new RESULT { Successful = false, Exception = e.ToString() };
            }
        }

        /// <summary>
        /// Combines all docs from a list of byte arrays into one.
        /// </summary>
        /// <param name="docs">List containing docs in byte array form</param>
        /// <returns>Single byte array containing all the docs in one doc, ready to be written to file.</returns>
        public static byte[] CombineDocs(IList<byte[]> docs)
        {
            MemoryStream mainStream = new MemoryStream();

            mainStream.Write(docs[0], 0, docs[0].Length);
            mainStream.Position = 0;

            int pointer = 1;
            byte[] ret;
            try
            {
                using (WordprocessingDocument mainDocument = WordprocessingDocument.Open(mainStream, true))
                {

                    XElement newBody = XElement.Parse(mainDocument.MainDocumentPart.Document.Body.OuterXml);
                    newBody.Add(XElement.Parse(new Paragraph(new Run(new Break { Type = BreakValues.Page })).OuterXml));
                    for (pointer = 1; pointer < docs.Count; pointer++)
                    {
                        WordprocessingDocument tempDocument = WordprocessingDocument.Open(new MemoryStream(docs[pointer]), true);
                        XElement tempBody = XElement.Parse(tempDocument.MainDocumentPart.Document.Body.OuterXml);

                        newBody.Add(tempBody);
                        if (pointer != docs.Count - 1)
                            newBody.Add(XElement.Parse(new Paragraph(new Run(new Break { Type = BreakValues.Page })).OuterXml));

                        mainDocument.MainDocumentPart.Document.Body = new Body(newBody.ToString());
                        mainDocument.MainDocumentPart.Document.Save();
                        mainDocument.Package.Flush();
                    }
                }
            }
            catch (OpenXmlPackageException oxmle)
            {
                throw new Exception(oxmle.ToString());
            }
            finally
            {
                ret = mainStream.ToArray();
                mainStream.Close();
                mainStream.Dispose();
            }
            return (ret);
        }
    }
}
