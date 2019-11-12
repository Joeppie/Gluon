using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Gluon
{
    static class OpenXMLExtensions
    {
        public static Run NewParaWithRun(this Body body, string text)
        {
            Paragraph p = new Paragraph();
            Run r = new Run();
            parseTextForOpenXML(r, text);
            p.Append(r);
            body.Append(p);
            return r;
        }

        public static Run Bold(this Run run)
        {
            RunProperties properties = run.RunProperties ?? new RunProperties();
            {
                properties.Append(new Bold());
            }
            run.RunProperties = properties;
            return run;
        }

        public static Run Size(this Run run, int size)
        {
            RunProperties properties = run.RunProperties ?? new RunProperties();
            {
                properties.Append(new FontSize() { Val = size.ToString() });
            }
            run.RunProperties = properties;
            return run;
        }


        public static Run Font(this Run run, string familyName)
        {
            RunProperties properties = run.RunProperties ?? new RunProperties();
            {
                properties.Append(new RunFonts { Ascii = familyName });
            }
            run.RunProperties = properties;
            return run;
        }

        public static Run Highlight(this Run run, HighlightColorValues color = HighlightColorValues.Red)
        {
            RunProperties properties = run.RunProperties ?? new RunProperties();
            {
                properties.Append(new Highlight { Val = color });
            }
            run.RunProperties = properties;
            return run;
        }



        public static void parseTextForOpenXML(Run run, string textualData, int fontSize = 8, bool bold = false)
        {
            string[] newLineArray = { "\r\n", "\n" };
            string[] textArray = textualData.Split(newLineArray, StringSplitOptions.None);

            bool first = true;

            foreach (string line in textArray)
            {
                if (!first)
                {
                    run.Append(new Break());
                }

                first = false;

                Text txt = new Text();
                txt.Text = line;
                run.Append(txt);
            }


        }
    }
}
