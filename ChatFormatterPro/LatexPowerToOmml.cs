// LatexPowerToOmml.cs
// STATUS: OBSOLETE — do not use in the math pipeline.
//
// WHY:
//   This file contains a separate LaTeX parser that produces XML strings
//   (raw string concatenation). DocxExporter already contains a complete,
//   superior OMML builder that uses the strongly-typed DocumentFormat.OpenXml SDK.
//   Two parsers for the same job creates maintenance confusion and potential
//   for mismatched output if one is wired in accidentally.
//
//   Additionally, this parser's XML string output cannot be safely mixed with
//   DocumentFormat.OpenXml SDK objects without an XmlElement round-trip,
//   which is fragile and unnecessary.
//
// ACTION:
//   Keep this file in your project so existing references compile,
//   but do NOT call LatexPowerToOmml.ToOmml() from anywhere.
//   All OMML generation goes through DocxExporter.BuildOfficeMath().
//
// FUTURE:
//   If you want to expand the OMML parser, extend DocxExporter's
//   ParseSequence() method — do not extend this file.

using System;
using System.Collections.Generic;
using System.Text;

namespace ChatFormatterPro
{
    /// <summary>
    /// OBSOLETE. Do not use. See file header for explanation.
    /// All OMML generation is handled by DocxExporter.BuildOfficeMath().
    /// </summary>
    [Obsolete("Use DocxExporter.BuildOfficeMath() instead. See file header.")]
    public static class LatexPowerToOmml
    {
        private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        [Obsolete("Do not call. See class-level Obsolete notice.")]
        public static string ToOmml(string latex)
        {
            throw new NotSupportedException(
                "LatexPowerToOmml is obsolete. Use DocxExporter.BuildOfficeMath() instead.");
        }
    }
}
