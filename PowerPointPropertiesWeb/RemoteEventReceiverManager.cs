using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SystemIO = System.IO;
using DocumentFormat.OpenXml.Drawing;
using System.Text.RegularExpressions;

namespace PowerPointPropertiesWeb
{
    public class RemoteEventReceiverManager
    {
        private const string RECEIVER_NAME_ADDED = "PowerPointPropertiesWeb";
        private const string LIST_TITLE = "documents";
        private bool saveFile = false;

        public void AssociateRemoteEventsToHostWeb(ClientContext clientContext)
        {
            var rerList = clientContext.Web.Lists.GetByTitle(LIST_TITLE);
            clientContext.Load(rerList);
            clientContext.ExecuteQuery();

            bool rerExists = false;
            if (!rerExists)
            {
                OperationContext op = OperationContext.Current;
                Message msg = op.RequestContext.RequestMessage;

                EventReceiverDefinitionCreationInformation receiver = new EventReceiverDefinitionCreationInformation();

                //receiver.EventType = EventReceiverType.ItemAdded;
                //receiver.ReceiverUrl = msg.Headers.To.ToString();
                //receiver.ReceiverName = EventReceiverType.ItemAdded.ToString();
                //receiver.Synchronization = EventReceiverSynchronization.Synchronous;
                //receiver.SequenceNumber = 1000;
                //rerList.EventReceivers.Add(receiver);
                //clientContext.ExecuteQuery();
                //System.Diagnostics.Trace.WriteLine("Added ItemAdded receiver at " + receiver.ReceiverUrl);

                receiver = new EventReceiverDefinitionCreationInformation();
                receiver.EventType = EventReceiverType.ItemUpdated;
                receiver.ReceiverUrl = msg.Headers.To.ToString();
                receiver.ReceiverName = EventReceiverType.ItemUpdated.ToString();
                receiver.Synchronization = EventReceiverSynchronization.Synchronous;
                receiver.SequenceNumber = 1000;
                rerList.EventReceivers.Add(receiver);
                clientContext.ExecuteQuery();

                System.Diagnostics.Trace.WriteLine("Added ItemUpdated receiver at " + receiver.ReceiverUrl);
            }
        }

        public void RemoveEventReceiversFromHostWeb(ClientContext clientContext)
        {
            List myList = clientContext.Web.Lists.GetByTitle(LIST_TITLE);
            clientContext.Load(myList, p => p.EventReceivers);
            clientContext.ExecuteQuery();

            var rer = myList.EventReceivers.Where(e => e.ReceiverName == RECEIVER_NAME_ADDED).FirstOrDefault();

            try
            {
                System.Diagnostics.Trace.WriteLine("Removing receiver at "
                        + rer.ReceiverUrl);

                var rerList = myList.EventReceivers.Where(e => e.ReceiverUrl == rer.ReceiverUrl).ToList<EventReceiverDefinition>();

                foreach (var rerFromUrl in rerList)
                {
                    //This will fail when deploying via F5, but works
                    //when deployed to production
                    rerFromUrl.DeleteObject();
                }
                clientContext.ExecuteQuery();
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }

            clientContext.ExecuteQuery();
        }

        public void ItemUpdatedToListEventHandler(ClientContext clientContext, Guid listId, int listItemId)
        {
            try
            {
                UpdatePowerPointProperties(clientContext, listId, listItemId);
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }

        public void ItemAddedToListEventHandler(ClientContext clientContext, Guid listId, int listItemId)
        {
            try
            {
                UpdatePowerPointProperties(clientContext, listId, listItemId);
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }

        public Dictionary<SlidePart, SlideId> SlideDictionary { get; set; }
        public List<PowerPointParameter> PowerPointTokenReplacements { get; set; }

        void UpdatePowerPointProperties(ClientContext clientContext, Guid listId, int listItemId)
        {
            List list = clientContext.Web.Lists.GetById(listId);
            ListItem item = list.GetItemById(listItemId);
            clientContext.Load(item);
            clientContext.ExecuteQuery();

            if (item["File_x0020_Type"].ToString() != "pptx")
                return;

            File file = item.File;
            ClientResult<SystemIO.Stream> data = file.OpenBinaryStream();

            // Load the Stream data for the file
            Site site = clientContext.Site;
            clientContext.Load(site);
            clientContext.Load(file);
            clientContext.ExecuteQuery();

            List<FieldTitles> fieldTypes = new List<FieldTitles>();
            fieldTypes.Add(new FieldTitles("Author", FieldTitleType.UserLookup));
            fieldTypes.Add(new FieldTitles("BusinessConsultant", FieldTitleType.Lookup));
            fieldTypes.Add(new FieldTitles("BusinessConsultantFirstName", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("BusinessConsultantLastName", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("BusinessConsultantEmail", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("BusinessConsultantMobile", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("FunctionalConsultant", FieldTitleType.Lookup));
            fieldTypes.Add(new FieldTitles("FunctionalConsultantFirstName", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("FunctionalConsultantLastName", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("FunctionalConsultantEmail", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("FunctionalConsultantMobile", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("ProjectManager", FieldTitleType.Lookup));
            fieldTypes.Add(new FieldTitles("ProjectManagerFirstName", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("ProjectManagerLastName", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("ProjectManagerEmail", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("ProjectManagerMobile", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("ClientLeadTitle", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("ClientLeadFirstName", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("ClientLeadLastName", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("ClientLeadEmail", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("ClientLeadMobile", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("ProjectTitle", FieldTitleType.String));
            fieldTypes.Add(new FieldTitles("Client", FieldTitleType.String));
 
            PowerPointTokenReplacements = new List<PowerPointParameter>();
            foreach (FieldTitles fieldTitle in fieldTypes)
            {
                string value = "";
                if (item[fieldTitle.title] == null)
                    continue;

                switch (fieldTitle.type)
                {
                    case FieldTitleType.UserLookup:
                        FieldUserValue fuv = (FieldUserValue)item[fieldTitle.title];
                        value = fuv.LookupValue;
                        break;
                    case FieldTitleType.String:
                        value = item[fieldTitle.title].ToString();
                        break;
                    case FieldTitleType.Lookup:
                        FieldLookupValue lv = (FieldLookupValue)item[fieldTitle.title];
                        value = lv.LookupValue;
                        break;
                    default:
                        value = item[fieldTitle.title].ToString();
                        break;
                }
                if (string.IsNullOrEmpty(value)) continue;

                PowerPointTokenReplacements.Add(new PowerPointParameter() { Name = "[#" + fieldTitle.title + "#]", Text = value });
            }

            if (!PowerPointTokenReplacements.Any()) return;

            using (SystemIO.MemoryStream memoryStream = new SystemIO.MemoryStream())
            {
                data.Value.CopyTo(memoryStream);
                memoryStream.Seek(0, SystemIO.SeekOrigin.Begin);

                using (PresentationDocument doc = PresentationDocument.Open(memoryStream, true))
                {
                    // Get the presentation part from the presentation document.
                    var presentationPart = doc.PresentationPart;

                    // Get the presentation from the presentation part.
                    var presentation = presentationPart.Presentation;

                    var slideList = new List<SlidePart>();
                    SlideDictionary = new Dictionary<SlidePart, SlideId>();

                    //get available slide list
                    foreach (SlideId slideID in presentation.SlideIdList)
                    {
                        var slide = (SlidePart)presentationPart.GetPartById(slideID.RelationshipId);
                        slideList.Add(slide);
                        SlideDictionary.Add(slide, slideID);//add to dictionary to be used when needed
                    }

                    //loop all slides and replace images and texts
                    foreach (var slide in slideList)
                    {
                        //ReplaceImages(presentationDocument, slide); //replace images by name

                        var paragraphs = slide.Slide.Descendants<Paragraph>().ToList(); //get all paragraphs in the slide

                        foreach (var paragraph in paragraphs)
                        {
                            ReplaceText(paragraph); //replace text by placeholder name
                        }
                    }

                    var slideCount = presentation.SlideIdList.ToList().Count; //count slides
                    //DeleteSlide(presentation, slideList[slideCount - 1]); //delete last slide

                    presentation.Save(); //save document changes we've made
                }

                if (saveFile)
                {
                    // Seek to beginning before writing to the SharePoint server.
                    memoryStream.Seek(0, SystemIO.SeekOrigin.Begin);
                    FileCreationInformation fci = new FileCreationInformation();
                    fci.ContentStream = memoryStream;
                    fci.Overwrite = true;
                    fci.Url = "https://rapidcircle1com.sharepoint.com" + (string)item["FileRef"];

                    File uploadFile = list.RootFolder.Files.Add(fci);
                    clientContext.ExecuteQuery();
                }
            }
        }

        void ReplaceText(Paragraph paragraph)
        {
            var parent = paragraph.Parent; //get parent element - to be used when removing placeholder
            var dataParam = new PowerPointParameter();

            if (ContainsParam(paragraph, ref dataParam)) //check if paragraph is on our parameter list
            {
                var param = CloneParaGraphWithStyles(paragraph, dataParam.Name, dataParam.Text); // create new param - preserve styles
                parent.InsertBefore(param, paragraph);//insert new element

                paragraph.Remove();//delete placeholder
                saveFile = true;
            }
        }

        public bool ContainsParam(Paragraph paragraph, ref PowerPointParameter dataParam)
        {
            foreach (var param in PowerPointTokenReplacements)
            {
                if (!string.IsNullOrEmpty(param.Name) && paragraph.InnerText.ToLower().Contains(param.Name.ToLower()))
                {
                    dataParam = param;
                    return true;
                }
            }

            return false;
        }

        public static Paragraph CloneParaGraphWithStyles(Paragraph sourceParagraph, string paramKey, string text)
        {
            var xmlSource = sourceParagraph.OuterXml;

            Regex regex = new Regex("(\\[#[\\s\\S]*?#])", RegexOptions.CultureInvariant | RegexOptions.Compiled);
            string result = regex.Replace(xmlSource, text.Trim());

            //xmlSource = xmlSource.Replace(paramKey.Trim(), text.Trim());

            return new Paragraph(result);
        }

        bool FieldExists(Dictionary<string, object> FieldValues, string Field)
        {
            foreach (var field in FieldValues)
            {
                if (field.Key.Equals(Field))
                    return true;
            }
            return false;
        }

        static private void CopyStream(SystemIO.Stream source, SystemIO.Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;
            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }

    }

    public class FieldTitles
    {
        public string title;
        public FieldTitleType type;

        public FieldTitles(string Title, FieldTitleType Type)
        {
            title = Title;
            type = Type;
        }
    }

    public enum FieldTitleType
    {
        UserLookup,
        String,
        Lookup
    }

    public class LookupFieldSet
    {
        public string LookupField { get; set; }
        public List<LookupFieldMapping> FieldMappings { get; set; }

        public LookupFieldSet()
        {
            FieldMappings = new List<LookupFieldMapping>();
        }

    }

    public class LookupFieldMapping
    {
        public string LocalLookupField { get; set; }
        public string ParentSourceField { get; set; }

        public LookupFieldMapping(string localLookupField, string parentSourceField)
        {
            LocalLookupField = localLookupField;
            ParentSourceField = parentSourceField;
        }
    }

    public static class StringExtensions
    {
        public static string Replace(this string originalString, string oldValue, string newValue, StringComparison comparisonType)
        {
            int startIndex = 0;
            while (true)
            {
                startIndex = originalString.IndexOf(oldValue, startIndex, comparisonType);
                if (startIndex == -1)
                    break;

                originalString = originalString.Substring(0, startIndex) + newValue + originalString.Substring(startIndex + oldValue.Length);

                startIndex += newValue.Length;
            }

            return originalString;
        }

    }

    public class PowerPointParameter
    {
        public string Name { get; set; }
        public string Text { get; set; }
    }
}