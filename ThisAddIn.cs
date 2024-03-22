using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace Latin_Cyrillic_TextConvert
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // no code
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // no code
        }

        Word.Selection wordSelection;
        public void L_to_C_Text_Convert()
        {
            wordSelection = this.Application.Selection;
            if (wordSelection != null && wordSelection.Range != null)
            {
                wordSelection.Text = ConvertToCyrillic(wordSelection);
            }
        }

        public void C_to_L_Text_Convert()
        {
            Word.Selection wordSelection = this.Application.Selection;
            if (wordSelection != null && wordSelection.Range != null)
            {
                wordSelection.Text = ConvertToLatin(wordSelection);
            }
        }

        private string ConvertToCyrillic(Word.Selection SourceText)
        {
            string TempText = SourceText.Text, Converted = "";
            int l = TempText.Length;
            for (int i = 0; i < l; i++)
            {
                //----------------------------------------
                //lowercase
                if (TempText[i] == 'd' && TempText[i + 1] == 'ž')
                {
                    Converted += "џ";
                    i++;
                    continue;
                }
                if (TempText[i] == 'l' && TempText[i + 1] == 'j')
                {
                    Converted += "љ";
                    i++;
                    continue;
                }
                if (TempText[i] == 'n' && TempText[i + 1] == 'j')
                {
                    Converted += "њ";
                    i++;
                    continue;
                }
                //lowercase
                //----------------------------------------

                //----------------------------------------
                //uppercase
                if (TempText[i] == 'D' && TempText[i + 1] == 'ž')
                {
                    Converted += "Џ";
                    i++;
                    continue;
                }

                if (TempText[i] == 'L' && TempText[i + 1] == 'j')
                {
                    Converted += "Љ";
                    i++;
                    continue;
                }

                if (TempText[i] == 'N' && TempText[i + 1] == 'j')
                {
                    Converted += "Њ";
                    i++;
                    continue;
                }
                //uppercase
                //----------------------------------------

                switch (TempText[i])
                {
                    //----------------------------------------
                    //lowercase
                    case 'a': { Converted += "а"; break; }
                    case 'b': { Converted += "б"; break; }
                    case 'v': { Converted += "в"; break; }
                    case 'g': { Converted += "г"; break; }
                    case 'd': { Converted += "д"; break; }
                    case 'đ': { Converted += "ђ"; break; }
                    case 'e': { Converted += "е"; break; }
                    case 'ž': { Converted += "ж"; break; }
                    case 'z': { Converted += "з"; break; }
                    case 'i': { Converted += "и"; break; }
                    case 'j': { Converted += "ј"; break; }
                    case 'k': { Converted += "к"; break; }
                    case 'l': { Converted += "л"; break; }
                    case 'm': { Converted += "м"; break; }
                    case 'n': { Converted += "н"; break; }
                    case 'o': { Converted += "о"; break; }
                    case 'p': { Converted += "п"; break; }
                    case 'r': { Converted += "р"; break; }
                    case 's': { Converted += "с"; break; }
                    case 't': { Converted += "т"; break; }
                    case 'ć': { Converted += "ћ"; break; }
                    case 'u': { Converted += "у"; break; }
                    case 'f': { Converted += "ф"; break; }
                    case 'h': { Converted += "х"; break; }
                    case 'c': { Converted += "ц"; break; }
                    case 'č': { Converted += "ч"; break; }
                    case 'š': { Converted += "ш"; break; }
                    //lowercase
                    //----------------------------------------

                    //----------------------------------------
                    //uppercase
                    case 'A': { Converted += "А"; break; }
                    case 'B': { Converted += "Б"; break; }
                    case 'V': { Converted += "В"; break; }
                    case 'G': { Converted += "Г"; break; }
                    case 'D': { Converted += "Д"; break; }
                    case 'Đ': { Converted += "Ђ"; break; }
                    case 'E': { Converted += "Е"; break; }
                    case 'Ž': { Converted += "Ж"; break; }
                    case 'Z': { Converted += "З"; break; }
                    case 'I': { Converted += "И"; break; }
                    case 'J': { Converted += "Ј"; break; }
                    case 'K': { Converted += "К"; break; }
                    case 'L': { Converted += "Л"; break; }
                    case 'M': { Converted += "М"; break; }
                    case 'N': { Converted += "Н"; break; }
                    case 'O': { Converted += "О"; break; }
                    case 'P': { Converted += "П"; break; }
                    case 'R': { Converted += "Р"; break; }
                    case 'S': { Converted += "С"; break; }
                    case 'T': { Converted += "Т"; break; }
                    case 'Ć': { Converted += "Ћ"; break; }
                    case 'U': { Converted += "У"; break; }
                    case 'F': { Converted += "Ф"; break; }
                    case 'H': { Converted += "Х"; break; }
                    case 'C': { Converted += "Ц"; break; }
                    case 'Č': { Converted += "Ч"; break; }
                    case 'Š': { Converted += "Ш"; break; }
                    //uppercase
                    //----------------------------------------
                    default: { Converted += TempText[i]; break; }
                }
            }
            return Converted;
        }

        private string ConvertToLatin(Word.Selection SourceText)
        {
            string TempText = SourceText.Text, Converted = "";
            int l = TempText.Length;
            for (int i = 0; i < l; i++)
            {
                switch (TempText[i])
                {
                    //----------------------------------------
                    //lowercase
                    case 'а': { Converted += "а"; break; }
                    case 'б': { Converted += "b"; break; }
                    case 'в': { Converted += "v"; break; }
                    case 'г': { Converted += "g"; break; }
                    case 'д': { Converted += "d"; break; }
                    case 'ђ': { Converted += "đ"; break; }
                    case 'е': { Converted += "e"; break; }
                    case 'ж': { Converted += "ž"; break; }
                    case 'з': { Converted += "z"; break; }
                    case 'и': { Converted += "i"; break; }
                    case 'ј': { Converted += "j"; break; }
                    case 'к': { Converted += "k"; break; }
                    case 'л': { Converted += "l"; break; }
                    case 'љ': { Converted += "lj"; break; }
                    case 'м': { Converted += "m"; break; }
                    case 'н': { Converted += "n"; break; }
                    case 'њ': { Converted += "nj"; break; }
                    case 'о': { Converted += "o"; break; }
                    case 'п': { Converted += "p"; break; }
                    case 'р': { Converted += "r"; break; }
                    case 'с': { Converted += "s"; break; }
                    case 'т': { Converted += "t"; break; }
                    case 'ћ': { Converted += "ć"; break; }
                    case 'у': { Converted += "u"; break; }
                    case 'ф': { Converted += "f"; break; }
                    case 'х': { Converted += "h"; break; }
                    case 'ц': { Converted += "c"; break; }
                    case 'ч': { Converted += "č"; break; }
                    case 'џ': { Converted += "dž"; break; }
                    case 'ш': { Converted += "š"; break; }
                    //lowercase
                    //----------------------------------------

                    //----------------------------------------
                    //uppercase
                    case 'А': { Converted += "A"; break; }
                    case 'Б': { Converted += "B"; break; }
                    case 'В': { Converted += "V"; break; }
                    case 'Г': { Converted += "G"; break; }
                    case 'Д': { Converted += "D"; break; }
                    case 'Ђ': { Converted += "Đ"; break; }
                    case 'Е': { Converted += "E"; break; }
                    case 'Ж': { Converted += "Ž"; break; }
                    case 'З': { Converted += "Z"; break; }
                    case 'И': { Converted += "I"; break; }
                    case 'Ј': { Converted += "J"; break; }
                    case 'K': { Converted += "K"; break; }
                    case 'Л': { Converted += "L"; break; }
                    case 'Љ': { Converted += "Lj"; break; }
                    case 'М': { Converted += "M"; break; }
                    case 'Н': { Converted += "N"; break; }
                    case 'Њ': { Converted += "Nj"; break; }
                    case 'О': { Converted += "O"; break; }
                    case 'П': { Converted += "P"; break; }
                    case 'Р': { Converted += "R"; break; }
                    case 'С': { Converted += "S"; break; }
                    case 'Т': { Converted += "T"; break; }
                    case 'Ћ': { Converted += "Ć"; break; }
                    case 'У': { Converted += "U"; break; }
                    case 'Ф': { Converted += "F"; break; }
                    case 'Х': { Converted += "H"; break; }
                    case 'Ц': { Converted += "C"; break; }
                    case 'Ч': { Converted += "Č"; break; }
                    case 'Џ': { Converted += "Dž"; break; }
                    case 'Ш': { Converted += "Š"; break; }
                    //uppercase
                    //----------------------------------------
                    default: { Converted += TempText[i]; break; }
                }
            }
            return Converted;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
