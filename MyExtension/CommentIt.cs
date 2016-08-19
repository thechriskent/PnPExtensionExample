using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Framework.Provisioning.Providers;
using System.IO;
using System.Xml;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;

namespace MyExtension
{
    public class CommentIt : ITemplateProviderExtension
    {
        private string _comment;
        public void Initialize(object settings)
        {
            _comment = settings as string;
        }

        public OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate PostProcessGetTemplate(OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate template)
        {
            throw new NotImplementedException();
        }

        public System.IO.Stream PostProcessSaveTemplate(System.IO.Stream stream)
        {
            MemoryStream result = new MemoryStream();

            //Load up the Template Stream to an XmlDocument so that we can manipulate it directly
            XmlDocument doc = new XmlDocument();
            doc.Load(stream);
            XmlNamespaceManager nspMgr = new XmlNamespaceManager(doc.NameTable);
            nspMgr.AddNamespace("pnp", XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05);

            XmlNode root = doc.SelectSingleNode("//pnp:Provisioning", nspMgr);
            XmlNode commentNode = doc.CreateComment(_comment);
            root.PrependChild(commentNode);

            //Put it back into stream form for other provider extensions to have a go and to finish processing
            doc.Save(result);
            result.Position = 0;

            return (result);
        }

        public System.IO.Stream PreProcessGetTemplate(System.IO.Stream stream)
        {
            throw new NotImplementedException();
        }

        public OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate PreProcessSaveTemplate(OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate template)
        {
            throw new NotImplementedException();
        }

        public bool SupportsGetTemplatePostProcessing
        {
            get { return (false); }
        }

        public bool SupportsGetTemplatePreProcessing
        {
            get { return (false); }
        }

        public bool SupportsSaveTemplatePostProcessing
        {
            get { return (true); }
        }

        public bool SupportsSaveTemplatePreProcessing
        {
            get { return (false); }
        }
    }
}
