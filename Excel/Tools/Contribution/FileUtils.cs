﻿using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.ExcelApi.Tools.Contribution
{
    /// <summary>
    /// File related helper tools
    /// </summary>
    public class FileUtils
    {
        #region Fields

        private CommonUtils _owner;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        protected internal FileUtils(CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Get the current default file extension for a document type. The method is not aware of the MS Compatibilty pack in 2003 or below
        /// </summary>
        /// <param name="type">target document type</param>
        /// <returns>default extension for document type</returns>
        public string FileExtension(DocumentFormat type)
        {
            switch (type)
            {
                case DocumentFormat.Normal:
                    return _owner.ApplicationIs2007OrHigher ? "xlsx" : "xls";
                case DocumentFormat.Macros:
                    return _owner.ApplicationIs2007OrHigher ? "xlsm" : "xls";
                case DocumentFormat.Template:
                    return _owner.ApplicationIs2007OrHigher ? "xltx" : "xlt";
                case DocumentFormat.TemplateMacros:
                    return _owner.ApplicationIs2007OrHigher ? "xltm" : "xlt";
                case DocumentFormat.Binary:
                    return "xlb";
                case DocumentFormat.AddinMacros:
                    return _owner.ApplicationIs2007OrHigher ? "xlam" : "xla";
                default:
                    throw new ArgumentOutOfRangeException("type");
            }
        }

        /// <summary>
        /// Add dot extension to argument filename
        /// </summary>
        /// <param name="fileName">target file name</param>
        /// <param name="type">target document format</param> 
        /// <returns>filename with dot and extension</returns>
        public string Combine(string fileName, DocumentFormat type)
        {
            string dotSeperator = fileName.EndsWith(".", StringComparison.InvariantCultureIgnoreCase) ? String.Empty : ".";
            return System.IO.Path.Combine(fileName, dotSeperator + FileExtension(type));
        }

        /// <summary>
        /// Combines 2 arguments and document type to valid file path 
        /// </summary>
        /// <param name="directoryPath">target directory path</param>
        /// <param name="fileName">target file name</param>
        /// <param name="type">target document format</param>
        /// <returns>Combined file path</returns>
        public string Combine(string directoryPath, string fileName, DocumentFormat type)
        {
            string dotSeperator = fileName.EndsWith(".", StringComparison.InvariantCultureIgnoreCase) ? String.Empty : ".";
            return System.IO.Path.Combine(directoryPath, fileName + dotSeperator + FileExtension(type));
        }

        #endregion
    }
}
