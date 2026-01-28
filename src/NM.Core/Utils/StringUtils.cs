namespace NM.Core.Utils
{
    /// <summary>
    /// String utility functions ported from VBA modExport.bas
    /// </summary>
    public static class StringUtils
    {
        /// <summary>
        /// Returns the depth of a BOM item number by counting periods.
        /// "1" returns 1, "1.1" returns 2, "1.1.1" returns 3
        /// </summary>
        public static int AssemblyDepth(string itemNumber)
        {
            if (string.IsNullOrEmpty(itemNumber))
                return 0;

            int depth = 1;
            foreach (char c in itemNumber)
            {
                if (c == '.')
                    depth++;
            }
            return depth;
        }

        /// <summary>
        /// Extracts file name without extension from a full path.
        /// "c:\desktop\Part.sldprt" returns "Part"
        /// </summary>
        public static string FileNameWithoutExtension(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath))
                return string.Empty;

            int lastSlash = fullPath.LastIndexOf('\\');
            int lastDot = fullPath.LastIndexOf('.');

            if (lastSlash < 0) lastSlash = -1;
            if (lastDot < 0) lastDot = fullPath.Length;

            int start = lastSlash + 1;
            int length = lastDot - start;

            if (length <= 0)
                return fullPath.Substring(start);

            return fullPath.Substring(start, length);
        }

        /// <summary>
        /// Removes the instance suffix from a SolidWorks component name.
        /// "Part-1" returns "Part", "Assembly/Part-2" returns "Part"
        /// </summary>
        public static string RemoveInstance(string componentName)
        {
            if (string.IsNullOrEmpty(componentName))
                return string.Empty;

            // Remove branch prefix (everything before last /)
            int slashPos = componentName.LastIndexOf('/');
            if (slashPos >= 0)
                componentName = componentName.Substring(slashPos + 1);

            // Remove instance suffix (everything after last -)
            int dashPos = componentName.LastIndexOf('-');
            if (dashPos > 0)
                componentName = componentName.Substring(0, dashPos);

            return componentName;
        }

        /// <summary>
        /// Wraps a string in quotes for ERP export format.
        /// </summary>
        public static string QuoteMe(string value)
        {
            return "\"" + (value ?? string.Empty) + "\" ";
        }

        /// <summary>
        /// Pads an item number with leading zeros to ensure minimum length.
        /// </summary>
        public static string PadItemNumber(string subItemNumber, int minLength = 2)
        {
            if (string.IsNullOrEmpty(subItemNumber))
                return new string('0', minLength);

            while (subItemNumber.Length < minLength)
                subItemNumber = "0" + subItemNumber;

            return subItemNumber;
        }

        /// <summary>
        /// Gets the last segment of an item number (extension).
        /// "1.2.3" returns "3"
        /// </summary>
        public static string GetItemExtension(string itemNumber)
        {
            if (string.IsNullOrEmpty(itemNumber))
                return string.Empty;

            int dotPos = itemNumber.LastIndexOf('.');
            if (dotPos < 0)
                return itemNumber;

            return itemNumber.Substring(dotPos + 1);
        }

        /// <summary>
        /// Sanitizes a string for ERP export (removes problematic characters).
        /// </summary>
        public static string ParseString(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            // Remove tabs, newlines, and excessive whitespace
            return value.Replace("\t", " ").Replace("\r", "").Replace("\n", " ").Trim();
        }
    }
}
