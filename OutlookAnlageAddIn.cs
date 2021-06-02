using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OutlookAnlageAddIn
{
    /// <summary>
    /// This Add-In is Property of Konrad Werner!
    /// It was only written by Konrad Werner and is not created to modify!
    /// Contact: konradw01@outlook.de
    /// </summary>

    public partial class OutlookAnlageAddIn
    {
        InitialiseClass initialise = new InitialiseClass();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
                Application.ItemSend += new
                Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        /// <summary>
        /// Wenn der Senden Button gedrückt wird, wird überprüft ob das Mail Objekt korrekt ist,
        /// sowie ob schon Anlagen vorhanden sind. Wenn beides den Anforderungen entspricht,
        /// kann man eine Anlage hinzufügen.
        /// </summary>
        /// <param name="Item">
        /// Das gesendete Element.
        /// </param>
        /// <param name="Cancel">
        /// False, wenn das Ereignis auftritt. Falls dieses Argument durch die Ereignisprozedur auf True festlegt wird,
        /// wird das Schließen abgebrochen, und die Arbeitsmappe bleibt geöffnet.
        /// </param>
        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            Outlook.MailItem mail = Item as Outlook.MailItem;
            if (mail != null)
            {
                if (SearchContains(mail) && mail.Attachments.Count == 0)
                {
                    DialogResult result = MessageBox.Show(initialise.InitialiseMessageBoxText(), initialise.InitialiseMessageBoxCaption(), MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        string attachmentPath = AttachmentSearch();
                        if (attachmentPath != null)
                        {
                            mail.Attachments.Add(attachmentPath);
                            Cancel = true;
                        }
                        else
                        {
                            MessageBox.Show("No file selected!");
                            Cancel = true;
                        }
                    }
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        #region private Methods

        /// <summary>
        /// Überprüft den Inhalt der Mail nach Wörter aus der <see cref="searchList"/>.
        /// </summary>
        /// <param name="mail">
        /// Das aktuelle <see cref="Outlook.MailItem"/>-Objekt.
        /// </param>
        /// <returns>
        /// True, wenn ein Eintrag aus <see cref="searchList"/> gefunden wurde.
        /// False, wenn kein Eintrag aus <see cref="searchList"/> gefunden wurde.
        /// </returns>
        private bool SearchContains(Outlook.MailItem mail)
        {
            List<string> searchList = initialise.InitialiseList();
            for (int i = 0; i < searchList.Count; i++)
            {
                if (mail.Body.Contains(searchList[i]))
                {
                    return true;
                }
            }
            return false;
        }
        
        /// <summary>
        /// Ruft einen <see cref="FolderBrowserDialog"/> auf um die anzuhängende Datei zu suchen.
        /// </summary>
        /// <returns>
        /// Den Pfad der ausgewählten Datei.
        /// </returns>
        private string AttachmentSearch()
        {
            FileDialog fileDialog = new OpenFileDialog();
            //fileDialog. = "Select the File, you want to attach to your Mail!";
            DialogResult browseResult = fileDialog.ShowDialog();
            if (browseResult == DialogResult.OK)
            {
                return fileDialog.FileName;
            }
            return null;
        }

        #endregion Private Methods
                
    }
}
