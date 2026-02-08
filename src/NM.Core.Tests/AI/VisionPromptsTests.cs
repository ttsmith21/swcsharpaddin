using Xunit;
using NM.Core.AI;
using NM.Core.AI.Models;

namespace NM.Core.Tests.AI
{
    public class VisionPromptsTests
    {
        [Fact]
        public void TitleBlockPrompt_ContainsRequiredFields()
        {
            string prompt = VisionPrompts.GetTitleBlockPrompt();

            Assert.Contains("part_number", prompt);
            Assert.Contains("description", prompt);
            Assert.Contains("revision", prompt);
            Assert.Contains("material", prompt);
            Assert.Contains("finish", prompt);
            Assert.Contains("drawn_by", prompt);
        }

        [Fact]
        public void TitleBlockPrompt_RequestsJsonOnly()
        {
            string prompt = VisionPrompts.GetTitleBlockPrompt();
            Assert.Contains("ONLY valid JSON", prompt);
        }

        [Fact]
        public void FullPagePrompt_ContainsAllSections()
        {
            string prompt = VisionPrompts.GetFullPagePrompt();

            Assert.Contains("title_block", prompt);
            Assert.Contains("dimensions", prompt);
            Assert.Contains("manufacturing_notes", prompt);
            Assert.Contains("gdt_callouts", prompt);
            Assert.Contains("holes", prompt);
            Assert.Contains("bend_info", prompt);
            Assert.Contains("special_requirements", prompt);
        }

        [Fact]
        public void FullPagePrompt_ContainsNoteCategories()
        {
            string prompt = VisionPrompts.GetFullPagePrompt();

            Assert.Contains("deburr", prompt);
            Assert.Contains("finish", prompt);
            Assert.Contains("heat_treat", prompt);
            Assert.Contains("weld", prompt);
            Assert.Contains("machine", prompt);
            Assert.Contains("inspect", prompt);
        }

        [Fact]
        public void ContextPrompt_IncludesKnownData()
        {
            var context = new VisionContext
            {
                KnownPartNumber = "TEST-123",
                KnownMaterial = "304 SS",
                KnownThickness_in = 0.125
            };

            string prompt = VisionPrompts.GetFullPagePromptWithContext(context);

            Assert.Contains("TEST-123", prompt);
            Assert.Contains("304 SS", prompt);
            Assert.Contains("0.1250", prompt);
            Assert.Contains("discrepancies", prompt);
        }

        [Fact]
        public void ContextPrompt_WithoutContext_ReturnsSameAsBase()
        {
            string basePrompt = VisionPrompts.GetFullPagePrompt();
            string contextPrompt = VisionPrompts.GetFullPagePromptWithContext(null);
            Assert.Equal(basePrompt, contextPrompt);
        }

        [Fact]
        public void ContextPrompt_EmptyContext_ReturnsSameAsBase()
        {
            string basePrompt = VisionPrompts.GetFullPagePrompt();
            string contextPrompt = VisionPrompts.GetFullPagePromptWithContext(new VisionContext());
            Assert.Equal(basePrompt, contextPrompt);
        }

        [Fact]
        public void SystemPrompt_IsNotEmpty()
        {
            string prompt = VisionPrompts.GetSystemPrompt();
            Assert.False(string.IsNullOrWhiteSpace(prompt));
            Assert.Contains("manufacturing engineer", prompt);
        }
    }
}
