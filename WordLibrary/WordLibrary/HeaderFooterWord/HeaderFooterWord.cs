using Xceed.Document.NET;
using Xceed.Words.NET;

namespace WordLibrary
{
    public class HeaderFooterWord
    {
        /// <summary>
        /// Ajoutez trois types différents d'en-têtes et de pieds de page à un document.
        /// </summary>
        public static void HeadersFooters(DocX document, string title)
        {
            // Insert a Paragraph in the first page of the document.
            var p1 = document.InsertParagraph("This is the ").Append("first").Bold().Append(" page Content.");
            p1.SpacingBefore(70d);
            p1.InsertPageBreakAfterSelf();

            // Insert a Paragraph in the second page of the document.
            var p2 = document.InsertParagraph("This is the ").Append("second").Bold().Append(" page Content.");
            p2.InsertPageBreakAfterSelf();

            // Insert a Paragraph in the third page of the document.
            var p3 = document.InsertParagraph("This is the ").Append("third").Bold().Append(" page Content.");
            p3.InsertPageBreakAfterSelf();

            // Insert a Paragraph in the third page of the document.
            var p4 = document.InsertParagraph("This is the ").Append("fourth").Bold().Append(" page Content.");

            // Add Headers and Footers to the document.
            document.AddHeaders();
            document.AddFooters();

            // Force the first page to have a different Header and Footer.
            document.DifferentFirstPage = true;

            // Force odd & even pages to have different Headers and Footers.
            document.DifferentOddAndEvenPages = true;

            // Insert a Paragraph into the first Header.
            document.Headers.First.InsertParagraph("This is the ").Append("first").Bold().Append(" page header");

            // Insert a Paragraph into the first Footer.
            document.Footers.First.InsertParagraph("This is the ").Append("first").Bold().Append(" page footer");

            // Insert a Paragraph into the even Header.
            document.Headers.Even.InsertParagraph("This is an ").Append("even").Bold().Append(" page header");

            // Insert a Paragraph into the even Footer.
            document.Footers.Even.InsertParagraph("This is an ").Append("even").Bold().Append(" page footer");

            // Insert a Paragraph into the odd Header.
            document.Headers.Odd.InsertParagraph("This is an ").Append("odd").Bold().Append(" page header");

            // Insert a Paragraph into the odd Footer.
            document.Footers.Odd.InsertParagraph("This is an ").Append("odd").Bold().Append(" page footer");

            // Add the page number in the first Footer.
            document.Footers.First.InsertParagraph("Page #").AppendPageNumber(PageNumberFormat.normal);

            // Add the page number in the even Footers.
            document.Footers.Even.InsertParagraph("Page #").AppendPageNumber(PageNumberFormat.normal);

            // Add the page number in the odd Footers.
            document.Footers.Odd.InsertParagraph("Page #").AppendPageNumber(PageNumberFormat.normal);

            document.Save();
        }
    }
}
