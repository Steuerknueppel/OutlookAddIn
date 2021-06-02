using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAnlageAddIn
{
    public class InitialiseClass
    {
        /// <summary>
        /// Füllt die <see cref="searchList"/> mit Einträgen.
        /// </summary>
        public List<string> InitialiseList()
        {
            List<string> searchList = new List<string>();
            searchList.Add("Anlage");
            searchList.Add("anlage");
            searchList.Add("Attachment");
            searchList.Add("attachment");
            searchList.Add("Anbei");
            searchList.Add("anbei");
            searchList.Add("Enclosed");
            searchList.Add("enclosed");
            searchList.Add("Anhang");
            searchList.Add("anhang");
            searchList.Add("Angehangen");
            searchList.Add("angehangen");
            searchList.Add("Arrangement");
            searchList.Add("arrengement");
            searchList.Add("Attached");
            searchList.Add("attached");
            searchList.Add("Angehangen");
            searchList.Add("angehangen");

            return searchList;
        }

        /// <summary>
        /// Läd die Parameter der <see cref="MessageBox"/>.
        /// </summary>
        public string InitialiseMessageBoxText()
        {
            string text = "If you forgot to attach a file, u can do it now. Just click 'Yes' to attach a file or 'No' to Send the Mail!";
            return text;
        }

        public string InitialiseMessageBoxCaption()
        {
            string caption = "Forgot your Attachment?";
            return caption;
        }

    }
}
