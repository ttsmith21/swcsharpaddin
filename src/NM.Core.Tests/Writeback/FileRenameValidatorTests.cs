using System.IO;
using Xunit;
using NM.Core.Reconciliation.Models;
using NM.Core.Writeback;

namespace NM.Core.Tests.Writeback
{
    public class FileRenameValidatorTests
    {
        private readonly FileRenameValidator _validator = new FileRenameValidator();

        // --- SanitizeFileName ---

        [Fact]
        public void SanitizeFileName_RemovesInvalidChars()
        {
            string result = FileRenameValidator.SanitizeFileName("part<>:1234");
            Assert.Equal("part1234", result);
        }

        [Fact]
        public void SanitizeFileName_CollapsesMultipleSpaces()
        {
            string result = FileRenameValidator.SanitizeFileName("part  number   1234");
            Assert.Equal("part number 1234", result);
        }

        [Fact]
        public void SanitizeFileName_TrimsLeadingTrailingDots()
        {
            string result = FileRenameValidator.SanitizeFileName("..hidden..");
            Assert.Equal("hidden", result);
        }

        [Fact]
        public void SanitizeFileName_ReturnsNullForEmpty()
        {
            Assert.Null(FileRenameValidator.SanitizeFileName(""));
            Assert.Null(FileRenameValidator.SanitizeFileName(null));
            Assert.Null(FileRenameValidator.SanitizeFileName("   "));
        }

        [Fact]
        public void SanitizeFileName_PreservesValidChars()
        {
            string result = FileRenameValidator.SanitizeFileName("12345-01 Rev C");
            Assert.Equal("12345-01 Rev C", result);
        }

        // --- IsValidFileName ---

        [Fact]
        public void IsValidFileName_AcceptsNormalNames()
        {
            Assert.True(FileRenameValidator.IsValidFileName("12345-01.SLDPRT"));
            Assert.True(FileRenameValidator.IsValidFileName("bracket_rev_c.SLDPRT"));
        }

        [Fact]
        public void IsValidFileName_RejectsReservedNames()
        {
            Assert.False(FileRenameValidator.IsValidFileName("CON.SLDPRT"));
            Assert.False(FileRenameValidator.IsValidFileName("PRN.SLDPRT"));
            Assert.False(FileRenameValidator.IsValidFileName("NUL.SLDPRT"));
            Assert.False(FileRenameValidator.IsValidFileName("COM1.SLDPRT"));
            Assert.False(FileRenameValidator.IsValidFileName("LPT1.SLDPRT"));
        }

        [Fact]
        public void IsValidFileName_RejectsEmptyName()
        {
            Assert.False(FileRenameValidator.IsValidFileName(""));
            Assert.False(FileRenameValidator.IsValidFileName(null));
        }

        // --- GenerateNewPath ---

        [Fact]
        public void GenerateNewPath_PreservesExtensionAndDirectory()
        {
            string result = _validator.GenerateNewPath("12345-01",
                Path.Combine("C:", "Parts", "old-name.SLDPRT"));

            Assert.Equal(Path.Combine("C:", "Parts", "12345-01.SLDPRT"), result);
        }

        [Fact]
        public void GenerateNewPath_SanitizesPartNumber()
        {
            string result = _validator.GenerateNewPath("12345<>01",
                Path.Combine("C:", "Parts", "old.SLDPRT"));

            Assert.Equal(Path.Combine("C:", "Parts", "1234501.SLDPRT"), result);
        }

        [Fact]
        public void GenerateNewPath_ReturnsNullForEmptyPartNumber()
        {
            Assert.Null(_validator.GenerateNewPath("", "/some/path.SLDPRT"));
            Assert.Null(_validator.GenerateNewPath(null, "/some/path.SLDPRT"));
        }

        [Fact]
        public void GenerateNewPath_ReturnsNullForEmptyPath()
        {
            Assert.Null(_validator.GenerateNewPath("12345", ""));
            Assert.Null(_validator.GenerateNewPath("12345", null));
        }

        // --- GenerateDrawingPath ---

        [Fact]
        public void GenerateDrawingPath_ChangesExtension()
        {
            string result = _validator.GenerateDrawingPath(
                Path.Combine("C:", "Parts", "12345-01.SLDPRT"));

            Assert.Equal(Path.Combine("C:", "Parts", "12345-01.SLDDRW"), result);
        }

        [Fact]
        public void GenerateDrawingPath_ReturnsNullForEmpty()
        {
            Assert.Null(_validator.GenerateDrawingPath(""));
            Assert.Null(_validator.GenerateDrawingPath(null));
        }

        // --- Validate ---

        [Fact]
        public void Validate_RejectsNullSuggestion()
        {
            var result = _validator.Validate(null);
            Assert.False(result.IsValid);
            Assert.Contains("null", result.ErrorMessage);
        }

        [Fact]
        public void Validate_RejectsEmptyOldPath()
        {
            var rename = new RenameSuggestion { OldPath = "", NewPath = "/new.SLDPRT" };
            var result = _validator.Validate(rename);
            Assert.False(result.IsValid);
        }

        [Fact]
        public void Validate_RejectsEmptyNewPath()
        {
            var rename = new RenameSuggestion { OldPath = "/old.SLDPRT", NewPath = "" };
            var result = _validator.Validate(rename);
            Assert.False(result.IsValid);
        }

        [Fact]
        public void Validate_RejectsNonExistentSource()
        {
            var rename = new RenameSuggestion
            {
                OldPath = Path.Combine(Path.GetTempPath(), "nonexistent_12345.SLDPRT"),
                NewPath = Path.Combine(Path.GetTempPath(), "new_12345.SLDPRT")
            };

            var result = _validator.Validate(rename);
            Assert.False(result.IsValid);
            Assert.Contains("not found", result.ErrorMessage);
        }

        [Fact]
        public void Validate_ValidRename_WithExistingSourceFile()
        {
            // Create a temp file to act as the source
            string tempFile = Path.GetTempFileName();
            string newPath = tempFile + ".renamed";
            try
            {
                var rename = new RenameSuggestion
                {
                    OldPath = tempFile,
                    NewPath = newPath
                };

                var result = _validator.Validate(rename);
                Assert.True(result.IsValid);
                Assert.Equal(tempFile, result.OldPath);
                Assert.Equal(newPath, result.NewPath);
                Assert.False(result.HasAssemblyReferences);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public void Validate_WarnsAboutAffectedAssemblies()
        {
            string tempFile = Path.GetTempFileName();
            string newPath = tempFile + ".renamed";
            try
            {
                var rename = new RenameSuggestion
                {
                    OldPath = tempFile,
                    NewPath = newPath
                };
                rename.AffectedAssemblies.Add("Assembly1.SLDASM");
                rename.AffectedAssemblies.Add("Assembly2.SLDASM");

                var result = _validator.Validate(rename);
                Assert.True(result.IsValid);
                Assert.True(result.HasAssemblyReferences);
                Assert.Equal(2, result.AffectedAssemblyCount);
                Assert.Contains("2 assembly", result.Warning);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public void Validate_RejectsDestinationAlreadyExists()
        {
            string tempFile1 = Path.GetTempFileName();
            string tempFile2 = Path.GetTempFileName();
            try
            {
                var rename = new RenameSuggestion
                {
                    OldPath = tempFile1,
                    NewPath = tempFile2 // Both exist
                };

                var result = _validator.Validate(rename);
                Assert.False(result.IsValid);
                Assert.Contains("already exists", result.ErrorMessage);
            }
            finally
            {
                File.Delete(tempFile1);
                File.Delete(tempFile2);
            }
        }
    }
}
