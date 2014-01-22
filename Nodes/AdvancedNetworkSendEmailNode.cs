using System;
using System.ComponentModel.Composition;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Collections.Generic;

using VVVV.Core.Logging;
using VVVV.PluginInterfaces.V1;
using VVVV.PluginInterfaces.V2;
using System.Text;
using System.Text.RegularExpressions;

namespace VVVV.Nodes
{

    #region PluginInfo
    [PluginInfo(Name = "SendEmail",
                Category = "Network",
                Help = "Send an email via smtp",
                Author = "phlegma",
				Version = "Advanced",
                Credits = "vux",
                Tags = "email,Smtp",
                AutoEvaluate = true)]
    #endregion PluginInfo
    public class AdvancedNetworkSendEmailNode : IPluginEvaluate
    {

        [Input("Host", DefaultString = "smtp.googlemail.com")]
        ISpread<string> FPinInHost;

        [Input("Port", MinValue = 0, MaxValue = 49151, AsInt = true, DefaultValue=587)]
        IDiffSpread<int> FPinInPort;

        [Input("Username", DefaultString="User@gmail.com")]
        IDiffSpread<string> FPinInUsername;

        [Input("Password", DefaultString="password")]
        IDiffSpread<string> FPinInPassword;

        [Input("Use SSL", DefaultValue=1)]
        IDiffSpread<bool> FPinInSSL;

        [Input("From", DefaultString = "User@gmail.com")]
        IDiffSpread<string> FPinInFrom;

        [Input("From Name", DefaultString = "Real Name")]
        IDiffSpread<string> FPinInFromName;

        [Input("To", DefaultString = "User@gmail.com")]
        IDiffSpread<string> FPinInTo;

        [Input("To Name", DefaultString = "Real Name")]
        IDiffSpread<string> FPinInToName;

        [Input("Subject")]
        IDiffSpread<string> FPinInSubject;

        [Input("Message")]
        IDiffSpread<string> FPinInMessage;

        [Input("EmailEncoding", EnumName = "EmailEncoding")]
        IDiffSpread<EnumEntry> FPinInEmailEncoding;

        [Input("Accept Html")]
        IDiffSpread<bool> FPinInIsHtml;

        [Input("Attachment", StringType = StringType.Filename)]
        ISpread<ISpread<string>> FPinInAttachment;
    	
    	[Input("Use inline attachment")]
        IDiffSpread<bool> FUseInlineAttachment;

        [Output("Success")]
        ISpread<bool> FPinOutSuccess;

        [Input("Send", IsBang = true, IsSingle=true)]
        IDiffSpread<bool> FPinInDoSend;

        [Import()]
        ILogger FLogger;

        string FError = "";

        [ImportingConstructor]
        public AdvancedNetworkSendEmailNode()
        { 
            var s = new string[]{"Ansi","Ascii","UTF8", "UTF32","Unicode"};
            //Please rename your Enum Type to avoid 
            //numerous "MyDynamicEnum"s in the system
            EnumManager.UpdateEnum("EmailEncoding", "Ansi", s);  
        }

        #region Evaluate
        public void Evaluate(int SpreadMax)
        {
            if (FPinInDoSend.IsChanged)
            {
                SpreadMax = SpreadUtils.SpreadMax(FPinInAttachment, FPinInDoSend, FPinInEmailEncoding, 
                    FPinInFrom, FPinInFromName, FPinInHost, FPinInIsHtml, FPinInMessage, FPinInPassword, 
                    FPinInPort, FPinInSSL, FPinInSubject, FPinInTo, FPinInToName, FPinInUsername, FUseInlineAttachment);
                FPinOutSuccess.SliceCount = SpreadMax;
                for (int i = 0; i < SpreadMax; i++)
                {
                    if (FPinInDoSend[i])
                    {
                        string Username = FPinInUsername[i];
                        string Pwd = FPinInPassword[i];
                        if (Username == null) { Username = ""; }
                        if (Pwd == null) { Pwd = ""; }

                        SmtpClient EmailClient = new SmtpClient(FPinInHost[i], FPinInPort[i]);
                        EmailClient.EnableSsl = FPinInSSL[i];
                        EmailClient.SendCompleted += new SendCompletedEventHandler(EmailClient_SendCompleted);

                        if (Username.Length > 0 && Pwd.Length > 0)
                        {
                            NetworkCredential SMTPUserInfo = new NetworkCredential(Username, Pwd);
                            EmailClient.Credentials = SMTPUserInfo;
                        }

                        FPinOutSuccess[i] = false;

                        try
                        {
                            string Message = FPinInMessage[i];
                            string Subject = FPinInSubject[i];

                            MailAddress fromAddress = new MailAddress(FPinInFrom[i], FPinInFromName[i]);
                            MailAddress toAdrress = new MailAddress(FPinInTo[i], FPinInToName[i]);
                            MailMessage mail = new MailMessage(fromAddress, toAdrress);
                        	
                            //Convert the Incomming Message to the corresponding encoding
                            switch (FPinInEmailEncoding[i].Index)
                            {
                                case (0):
                                    mail.BodyEncoding = mail.SubjectEncoding = Encoding.Default;
                                    break;
                                case (1):
                                    mail.BodyEncoding = mail.SubjectEncoding = Encoding.ASCII;
                                    break;
                                case (2):
                                    mail.BodyEncoding = mail.SubjectEncoding = Encoding.UTF8;
                                    break;
                                case (3):
                                    mail.BodyEncoding = mail.SubjectEncoding = Encoding.UTF32;
                                    break;
                                case (4):
                                    mail.BodyEncoding = mail.SubjectEncoding = Encoding.Unicode;
                                    break;
                                default:
                                    mail.BodyEncoding = mail.SubjectEncoding = Encoding.Default;
                                    break;
                            }

                        	mail.Subject = Subject;
                    		mail.IsBodyHtml = FPinInIsHtml[i];
                        	mail.Body = Message;
                        	
                            foreach (var filename in FPinInAttachment[i])
                            {
                                if (File.Exists(filename))
                                {
                                	if( !FUseInlineAttachment[i] )
                                	{
	                                    Attachment Attachment = new Attachment(filename);
	                                    ContentDisposition Disposition = Attachment.ContentDisposition;
	                                    Disposition.Inline = false;
	                                    mail.Attachments.Add(Attachment);
                                	}else{
                                		Attachment inline = new Attachment(filename);
                                		inline.ContentDisposition.Inline = true;
                                		inline.ContentDisposition.DispositionType = DispositionTypeNames.Inline;
                                		inline.ContentId = Path.GetFileName(filename);
                                		inline.ContentType.MediaType = MIMEAssistant.GetMIMEType(filename);
                                		inline.ContentType.Name = Path.GetFileName(filename);
                                		mail.Attachments.Add(inline);
                                	}
                                }
                            }
                        	
                            EmailClient.SendAsync(mail, i);
                        }
                        catch (Exception ex)
                        {
                            FLogger.Log(LogType.Error, ex.Message);
                        }
                        finally
                        {
                            EmailClient = null;
                        }
                    }
                }
            }

            if (!String.IsNullOrEmpty(FError))
            {
                FLogger.Log(LogType.Debug, FError);
                FError = "";
            }
        }

        void EmailClient_SendCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            // Get the unique identifier for this asynchronous operation.
            int index = (int)e.UserState;

            if (e.Cancelled)
            {
                FError = "Cancelt";
            }
            if (e.Error != null)
            {
                FError = e.Error.ToString();
            }
            else
            {
            	FPinOutSuccess[index] = true;
                FError = "Message sent.";
            }
        }
        #endregion

    }
	
	public static class MIMEAssistant
	{
	  private static readonly Dictionary<string, string> MIMETypesDictionary = new Dictionary<string, string>
	  {
	    {"ai", "application/postscript"},
	    {"aif", "audio/x-aiff"},
	    {"aifc", "audio/x-aiff"},
	    {"aiff", "audio/x-aiff"},
	    {"asc", "text/plain"},
	    {"atom", "application/atom+xml"},
	    {"au", "audio/basic"},
	    {"avi", "video/x-msvideo"},
	    {"bcpio", "application/x-bcpio"},
	    {"bin", "application/octet-stream"},
	    {"bmp", "image/bmp"},
	    {"cdf", "application/x-netcdf"},
	    {"cgm", "image/cgm"},
	    {"class", "application/octet-stream"},
	    {"cpio", "application/x-cpio"},
	    {"cpt", "application/mac-compactpro"},
	    {"csh", "application/x-csh"},
	    {"css", "text/css"},
	    {"dcr", "application/x-director"},
	    {"dif", "video/x-dv"},
	    {"dir", "application/x-director"},
	    {"djv", "image/vnd.djvu"},
	    {"djvu", "image/vnd.djvu"},
	    {"dll", "application/octet-stream"},
	    {"dmg", "application/octet-stream"},
	    {"dms", "application/octet-stream"},
	    {"doc", "application/msword"},
	    {"docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
	    {"dotx", "application/vnd.openxmlformats-officedocument.wordprocessingml.template"},
	    {"docm","application/vnd.ms-word.document.macroEnabled.12"},
	    {"dotm","application/vnd.ms-word.template.macroEnabled.12"},
	    {"dtd", "application/xml-dtd"},
	    {"dv", "video/x-dv"},
	    {"dvi", "application/x-dvi"},
	    {"dxr", "application/x-director"},
	    {"eps", "application/postscript"},
	    {"etx", "text/x-setext"},
	    {"exe", "application/octet-stream"},
	    {"ez", "application/andrew-inset"},
	    {"gif", "image/gif"},
	    {"gram", "application/srgs"},
	    {"grxml", "application/srgs+xml"},
	    {"gtar", "application/x-gtar"},
	    {"hdf", "application/x-hdf"},
	    {"hqx", "application/mac-binhex40"},
	    {"htm", "text/html"},
	    {"html", "text/html"},
	    {"ice", "x-conference/x-cooltalk"},
	    {"ico", "image/x-icon"},
	    {"ics", "text/calendar"},
	    {"ief", "image/ief"},
	    {"ifb", "text/calendar"},
	    {"iges", "model/iges"},
	    {"igs", "model/iges"},
	    {"jnlp", "application/x-java-jnlp-file"},
	    {"jp2", "image/jp2"},
	    {"jpe", "image/jpeg"},
	    {"jpeg", "image/jpeg"},
	    {"jpg", "image/jpeg"},
	    {"js", "application/x-javascript"},
	    {"kar", "audio/midi"},
	    {"latex", "application/x-latex"},
	    {"lha", "application/octet-stream"},
	    {"lzh", "application/octet-stream"},
	    {"m3u", "audio/x-mpegurl"},
	    {"m4a", "audio/mp4a-latm"},
	    {"m4b", "audio/mp4a-latm"},
	    {"m4p", "audio/mp4a-latm"},
	    {"m4u", "video/vnd.mpegurl"},
	    {"m4v", "video/x-m4v"},
	    {"mac", "image/x-macpaint"},
	    {"man", "application/x-troff-man"},
	    {"mathml", "application/mathml+xml"},
	    {"me", "application/x-troff-me"},
	    {"mesh", "model/mesh"},
	    {"mid", "audio/midi"},
	    {"midi", "audio/midi"},
	    {"mif", "application/vnd.mif"},
	    {"mov", "video/quicktime"},
	    {"movie", "video/x-sgi-movie"},
	    {"mp2", "audio/mpeg"},
	    {"mp3", "audio/mpeg"},
	    {"mp4", "video/mp4"},
	    {"mpe", "video/mpeg"},
	    {"mpeg", "video/mpeg"},
	    {"mpg", "video/mpeg"},
	    {"mpga", "audio/mpeg"},
	    {"ms", "application/x-troff-ms"},
	    {"msh", "model/mesh"},
	    {"mxu", "video/vnd.mpegurl"},
	    {"nc", "application/x-netcdf"},
	    {"oda", "application/oda"},
	    {"ogg", "application/ogg"},
	    {"pbm", "image/x-portable-bitmap"},
	    {"pct", "image/pict"},
	    {"pdb", "chemical/x-pdb"},
	    {"pdf", "application/pdf"},
	    {"pgm", "image/x-portable-graymap"},
	    {"pgn", "application/x-chess-pgn"},
	    {"pic", "image/pict"},
	    {"pict", "image/pict"},
	    {"png", "image/png"}, 
	    {"pnm", "image/x-portable-anymap"},
	    {"pnt", "image/x-macpaint"},
	    {"pntg", "image/x-macpaint"},
	    {"ppm", "image/x-portable-pixmap"},
	    {"ppt", "application/vnd.ms-powerpoint"},
	    {"pptx","application/vnd.openxmlformats-officedocument.presentationml.presentation"},
	    {"potx","application/vnd.openxmlformats-officedocument.presentationml.template"},
	    {"ppsx","application/vnd.openxmlformats-officedocument.presentationml.slideshow"},
	    {"ppam","application/vnd.ms-powerpoint.addin.macroEnabled.12"},
	    {"pptm","application/vnd.ms-powerpoint.presentation.macroEnabled.12"},
	    {"potm","application/vnd.ms-powerpoint.template.macroEnabled.12"},
	    {"ppsm","application/vnd.ms-powerpoint.slideshow.macroEnabled.12"},
	    {"ps", "application/postscript"},
	    {"qt", "video/quicktime"},
	    {"qti", "image/x-quicktime"},
	    {"qtif", "image/x-quicktime"},
	    {"ra", "audio/x-pn-realaudio"},
	    {"ram", "audio/x-pn-realaudio"},
	    {"ras", "image/x-cmu-raster"},
	    {"rdf", "application/rdf+xml"},
	    {"rgb", "image/x-rgb"},
	    {"rm", "application/vnd.rn-realmedia"},
	    {"roff", "application/x-troff"},
	    {"rtf", "text/rtf"},
	    {"rtx", "text/richtext"},
	    {"sgm", "text/sgml"},
	    {"sgml", "text/sgml"},
	    {"sh", "application/x-sh"},
	    {"shar", "application/x-shar"},
	    {"silo", "model/mesh"},
	    {"sit", "application/x-stuffit"},
	    {"skd", "application/x-koan"},
	    {"skm", "application/x-koan"},
	    {"skp", "application/x-koan"},
	    {"skt", "application/x-koan"},
	    {"smi", "application/smil"},
	    {"smil", "application/smil"},
	    {"snd", "audio/basic"},
	    {"so", "application/octet-stream"},
	    {"spl", "application/x-futuresplash"},
	    {"src", "application/x-wais-source"},
	    {"sv4cpio", "application/x-sv4cpio"},
	    {"sv4crc", "application/x-sv4crc"},
	    {"svg", "image/svg+xml"},
	    {"swf", "application/x-shockwave-flash"},
	    {"t", "application/x-troff"},
	    {"tar", "application/x-tar"},
	    {"tcl", "application/x-tcl"},
	    {"tex", "application/x-tex"},
	    {"texi", "application/x-texinfo"},
	    {"texinfo", "application/x-texinfo"},
	    {"tif", "image/tiff"},
	    {"tiff", "image/tiff"},
	    {"tr", "application/x-troff"},
	    {"tsv", "text/tab-separated-values"},
	    {"txt", "text/plain"},
	    {"ustar", "application/x-ustar"},
	    {"vcd", "application/x-cdlink"},
	    {"vrml", "model/vrml"},
	    {"vxml", "application/voicexml+xml"},
	    {"wav", "audio/x-wav"},
	    {"wbmp", "image/vnd.wap.wbmp"},
	    {"wbmxl", "application/vnd.wap.wbxml"},
	    {"wml", "text/vnd.wap.wml"},
	    {"wmlc", "application/vnd.wap.wmlc"},
	    {"wmls", "text/vnd.wap.wmlscript"},
	    {"wmlsc", "application/vnd.wap.wmlscriptc"},
	    {"wrl", "model/vrml"},
	    {"xbm", "image/x-xbitmap"},
	    {"xht", "application/xhtml+xml"},
	    {"xhtml", "application/xhtml+xml"},
	    {"xls", "application/vnd.ms-excel"},                        
	    {"xml", "application/xml"},
	    {"xpm", "image/x-xpixmap"},
	    {"xsl", "application/xml"},
	    {"xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
	    {"xltx","application/vnd.openxmlformats-officedocument.spreadsheetml.template"},
	    {"xlsm","application/vnd.ms-excel.sheet.macroEnabled.12"},
	    {"xltm","application/vnd.ms-excel.template.macroEnabled.12"},
	    {"xlam","application/vnd.ms-excel.addin.macroEnabled.12"},
	    {"xlsb","application/vnd.ms-excel.sheet.binary.macroEnabled.12"},
	    {"xslt", "application/xslt+xml"},
	    {"xul", "application/vnd.mozilla.xul+xml"},
	    {"xwd", "image/x-xwindowdump"},
	    {"xyz", "chemical/x-xyz"},
	    {"zip", "application/zip"}
	  };
	
	  public static string GetMIMEType(string fileName)
	  {
	    //get file extension
	    string extension = Path.GetExtension(fileName).ToLowerInvariant();
	
	    if (extension.Length > 0 && 
	        MIMETypesDictionary.ContainsKey(extension.Remove(0, 1)))
	    {
	      return MIMETypesDictionary[extension.Remove(0, 1)];
	    }
	    return "unknown/unknown";
	  }
	}
}