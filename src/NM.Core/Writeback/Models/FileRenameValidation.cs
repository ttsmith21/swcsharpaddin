namespace NM.Core.Writeback.Models
{
    /// <summary>
    /// Result of validating a file rename operation.
    /// </summary>
    public sealed class FileRenameValidation
    {
        /// <summary>True if the rename is safe to proceed.</summary>
        public bool IsValid { get; set; }

        /// <summary>Current file path.</summary>
        public string OldPath { get; set; }

        /// <summary>Proposed new file path.</summary>
        public string NewPath { get; set; }

        /// <summary>Error message if invalid, warning if valid but notable.</summary>
        public string Warning { get; set; }

        /// <summary>Error message when IsValid is false.</summary>
        public string ErrorMessage { get; set; }

        /// <summary>True if assemblies reference this file.</summary>
        public bool HasAssemblyReferences { get; set; }

        /// <summary>Number of assemblies that reference this file.</summary>
        public int AffectedAssemblyCount { get; set; }

        /// <summary>Creates an error result.</summary>
        public static FileRenameValidation Error(string message)
        {
            return new FileRenameValidation
            {
                IsValid = false,
                ErrorMessage = message
            };
        }

        /// <summary>Creates a skip result (no rename needed).</summary>
        public static FileRenameValidation Skip(string reason)
        {
            return new FileRenameValidation
            {
                IsValid = false,
                ErrorMessage = reason
            };
        }
    }
}
