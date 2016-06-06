using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HtmlConverter;
using System.IO;

namespace UnitTests
{
    [TestClass]
    public class ImageGenerator_Test
    {
        [TestMethod]
        public void Test_ImageGenerator_GetPngFromHtml()
        {
            #region Test HTML

            var html = "<html xmlns=\"http://www.w3.org/1999/xhtml\" lang=\"en-US\">\r\n\t<head>\r\n\t\t<title>RFQ &amp; Risk Assessment Review Meeting (APQP #2)</title>\r\n\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\r\n\t\t<meta name=\"created\" content=\"2016-06-01T17:32:00.0000000\" />\r\n\t</head>\r\n\t<body data-absolute-enabled=\"true\" style=\"font-family:Calibri;font-size:11pt\">\r\n\t\t<div style=\"position:absolute;left:48px;top:787px;width:457px\">\r\n\t\t\t<p style=\"margin-top:0pt;margin-bottom:0pt\"><span style=\"font-weight:bold\">Meeting Date: </span>6/1/2016 6:30 PM</p>\r\n\t\t\t<p style=\"margin-top:0pt;margin-bottom:0pt\"><span style=\"font-weight:bold\">Location: </span>Skype Meeting</p>\r\n\t\t\t<p style=\"margin-top:0pt;margin-bottom:0pt\"><sub style=\"font-size:1pt;color:#979797\">([10</sub></p>\r\n\t\t\t<p style=\"margin-top:0pt;margin-bottom:0pt\"><span style=\"font-weight:bold\">Participants</span></p>\r\n\t\t\t<p data-tag=\"discuss-with-person-a:completed\" style=\"margin-top:0pt;margin-bottom:0pt\"><img class=\"NoteTagImage\" role=\"checkbox\" aria-checked=\"true\" src=\"https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/94_16_M.16x16x32.png\" alt=\"In Attendance\" title=\"In Attendance\" style=\"left: -32px;\"><a href=\"mailto:ivan@neganov.com\">Ivan Neganov </a><span style=\"font-size:10pt;color:#7f7f7f\">(Meeting Organizer, Joined in Skype for Business)</span></p>\r\n\t\t\t<p data-tag=\"discuss-with-person-a\" style=\"margin-top:0pt;margin-bottom:0pt\"><img class=\"NoteTagImage\" role=\"checkbox\" aria-checked=\"false\" src=\"https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/94_16_N.16x16x32.png\" alt=\"In Attendance\" title=\"In Attendance\" style=\"left: -32px;\"><a href=\"mailto:sveta@neganov.com\">Sveta Semenenkova</a></p>\r\n\t\t\t<br />\r\n\t\t</div>\r\n\t\t<div style=\"position:absolute;left:48px;top:191px;width:852px\">\r\n\t\t\t<p style=\"margin-top:0pt;margin-bottom:0pt\"><span style=\"font-size:14pt;font-weight:bold\">Risks</span></p>\r\n\t\t\t<table style=\"border:1px solid;border-collapse:collapse\">\r\n\t\t\t\t<tr>\r\n\t\t\t\t\t<td style=\"border:1px solid\"><span style=\"font-weight:bold\">PE Risks:</span></td>\r\n\t\t\t\t\t<td style=\"border:1px solid\">\r\n                    <p data-tag=\"to-do\" style=\"margin-top:0pt;margin-bottom:0pt\"><img class=\"NoteTagImage\" role=\"checkbox\" aria-checked=\"false\" src=\"https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/3_16_N.16x16x32.png\" alt=\"To Do\" title=\"To Do\" style=\"left: -32px;\"><span style=\"color:black\">Need to understand how reduced output will affect part angularity over time.</span></p>\r\n\t\t\t\t\t<p data-tag=\"to-do\" style=\"margin-top:0pt;margin-bottom:0pt\"><img class=\"NoteTagImage\" role=\"checkbox\" aria-checked=\"false\" src=\"https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/3_16_N.16x16x32.png\" alt=\"To Do\" title=\"To Do\" style=\"left: -32px;\"><span style=\"color:black\">Need to confirm spring resonant frequency does not fall in engine operating range.</span></p>\r\n\t\t\t\t\t<p data-tag=\"to-do\" style=\"margin-top:0pt;margin-bottom:0pt\"><img class=\"NoteTagImage\" role=\"checkbox\" aria-checked=\"false\" src=\"https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/3_16_N.16x16x32.png\" alt=\"To Do\" title=\"To Do\" style=\"left: -32px;\"><span style=\"color:black\">Need to test how new plating performs on front plate \u2013 sees high thrust loads.</span></p>\r\n\t\t\t\t\t</td>\r\n\t\t\t\t</tr>\r\n\t\t\t\t<tr>\r\n\t\t\t\t\t<td style=\"border:1px solid\"><span style=\"font-weight:bold\">SMG Risks:</span></td>\r\n\t\t\t\t\t<td style=\"border:1px solid\">NO SMG RISKS</td>\r\n\t\t\t\t</tr>\r\n\t\t\t\t<tr>\r\n\t\t\t\t\t<td style=\"border:1px solid\"><span style=\"font-weight:bold\">MFE Risks:</span></td>\r\n\t\t\t\t\t<td style=\"border:1px solid\">\r\n                    <p data-tag=\"to-do:completed\" style=\"margin-top:0pt;margin-bottom:0pt\"><img class=\"NoteTagImage\" role=\"checkbox\" aria-checked=\"true\" src=\"https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/3_16_M.16x16x32.png\" alt=\"To Do\" title=\"To Do\" style=\"left: -32px;\"><span style=\"color:black\">Assumed that this part will take some volume from 600933, else there is not enough capacity on WC 90 and job would need extra labor.</span></p>\r\n\t\t\t\t\t<p data-tag=\"to-do:completed\" style=\"margin-top:0pt;margin-bottom:0pt\"><img class=\"NoteTagImage\" role=\"checkbox\" aria-checked=\"true\" src=\"https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/3_16_M.16x16x32.png\" alt=\"To Do\" title=\"To Do\" style=\"left: -32px;\"><span style=\"color:black\">Compression spring should have yellow dye added as a visual aid in identifying spring (avoid relying on scanning system to identify spring) </span></p>\r\n\t\t\t\t\t</td>\r\n\t\t\t\t</tr>\r\n\t\t\t\t<tr>\r\n\t\t\t\t\t<td style=\"border:1px solid\"><span style=\"font-weight:bold\">QA Risks:</span></td>\r\n\t\t\t\t\t<td style=\"border:1px solid\">\r\n                    <p data-tag=\"to-do:completed\" style=\"margin-top:0pt;margin-bottom:0pt\"><img class=\"NoteTagImage\" role=\"checkbox\" aria-checked=\"true\" src=\"https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/3_16_M.16x16x32.png\" alt=\"To Do\" title=\"To Do\" style=\"left: -32px;\"><span style=\"color:black\">More tolerance is needed for diamond characteristics. Current 600933 is running on PDN to open tolerance on parallelism.</span></p>\r\n\t\t\t\t\t<p data-tag=\"to-do:completed\" style=\"margin-top:0pt;margin-bottom:0pt\"><img class=\"NoteTagImage\" role=\"checkbox\" aria-checked=\"true\" src=\"https://s1-onenote-15.cdn.office.net:443/o/s/1670242202_resources/1033/NoteTags/3_16_M.16x16x32.png\" alt=\"To Do\" title=\"To Do\" style=\"left: -32px;\"><span style=\"color:black\">None of the diamond characteristics meet SPC requirements.</span></p>\r\n\t\t\t\t\t</td>\r\n\t\t\t\t</tr>\r\n\t\t\t</table>\r\n\t\t</div>\r\n\t\t<div style=\"position:absolute;left:48px;top:475px;width:850px\">\r\n\t\t\t<p style=\"margin-top:0pt;margin-bottom:0pt\"><span style=\"font-size:14pt;font-weight:bold\">Questions</span></p>\r\n\t\t\t<table style=\"border:1px solid;border-collapse:collapse\">\r\n\t\t\t\t<tr>\r\n\t\t\t\t\t<td style=\"border:1px solid\"><span style=\"font-weight:bold\">Can manufacturing produce the part as designed and quoted?:</span></td>\r\n\t\t\t\t\t<td style=\"border:1px solid\"><p style=\"margin-top:0pt;margin-bottom:0pt\">No, cannot meet SPC requirements.</p>\r\n\t\t\t\t\t<p style=\"margin-top:0pt;margin-bottom:0pt\">Load, centerline, and angularity do not meet SPC requirements.</p>\r\n\t\t\t\t\t<p style=\"margin-top:0pt;margin-bottom:0pt\">Damping is barely in spec.</p>\r\n\t\t\t\t\t</td>\r\n\t\t\t\t</tr>\r\n\t\t\t\t<tr>\r\n\t\t\t\t\t<td style=\"border:1px solid\"><span style=\"font-weight:bold\">Can the quoted suppliers produce the part as required? Has the supplier input been reviewed?:</span></td>\r\n\t\t\t\t\t<td style=\"border:1px solid\">Yes</td>\r\n\t\t\t\t</tr>\r\n\t\t\t\t<tr>\r\n\t\t\t\t\t<td style=\"border:1px solid\"><span style=\"font-weight:bold\">Are the quted suppliers (including sub tier (s)) that perform Special Processes (CQI  Controlled) approved?:</span></td>\r\n\t\t\t\t\t<td style=\"border:1px solid\"><p style=\"margin-top:0pt;margin-bottom:0pt\">Yes</p>\r\n\t\t\t\t\t<br />\r\n\t\t\t\t\t</td>\r\n\t\t\t\t</tr>\r\n\t\t\t\t<tr>\r\n\t\t\t\t\t<td style=\"border:1px solid\"><span style=\"font-weight:bold\">Can customer deadlines be achieved?:</span></td>\r\n\t\t\t\t\t<td style=\"border:1px solid\"><p style=\"margin-top:0pt;margin-bottom:0pt\">No. Laser machine will require 20 weeks. If no laser marking, then we can meet timing.</p>\r\n\t\t\t\t\t<br />\r\n\t\t\t\t\t</td>\r\n\t\t\t\t</tr>\r\n\t\t\t</table>\r\n\t\t</div>\r\n\t\t<div style=\"position:absolute;left:48px;top:115px;width:720px\">\r\n\t\t\t<p style=\"margin-top:0pt;margin-bottom:0pt\"><span style=\"font-size:14pt;color:#538135;font-weight:bold\">Meeting Details</span></p>\r\n\t\t\t<p style=\"margin-top:0pt;margin-bottom:0pt\"><span style=\"font-size:9pt\">Meeting Type: 1-2 \u2013 RFQ &amp; Risk Assessment Review Meeting (APQP#2)</span></p>\r\n\t\t</div>\r\n\t</body>\r\n</html>\r\n";

            #endregion

            string fileName = GetUnusedFileName(null);

            using (Stream stream = File.Open(
                fileName, 
                FileMode.OpenOrCreate, 
                FileAccess.Write))
            {
                ImageGenerator.GetPngFromHtml(stream, html);
            }
        }

        private string GetUnusedFileName(string name)
        {
            if(String.IsNullOrWhiteSpace(name))
            {
                name = "image.png";
            }

            if(File.Exists(name))
            {
                int dotIdx = name.IndexOf('.');
                int length = dotIdx - 5;
                int newNumber = 1;

                if (length > 0)
                {
                    string numberString = name.Substring(5, length);
                    newNumber = (Int32.Parse(numberString)) + 1;
                }

                string newName = "image" + newNumber + ".png";
                return GetUnusedFileName(newName);
            }
            else
            {
                return name;
            }
        }
    }
}
