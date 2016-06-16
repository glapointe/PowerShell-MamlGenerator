using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Text;
using System.Reflection;
using System.Xml;
using System.IO;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.PowerShell.MamlGenerator
{
    public class CmdletHelpGenerator
    {
        private static string _copyright = null;

        private static XmlTextWriter _writer = null;

        /*
        C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit "& {W:\Work\Lapointe.SharePoint.PowerShell\Lapointe.SharePoint.PowerShell\GenerateHelp.ps1 W:\Work\Lapointe.SharePoint.PowerShell\Lapointe.SharePoint.PowerShell\bin\Debug\}"
         */
        public static void GenerateHelp(string outputPath, bool oneFile)
        {
            //System.Diagnostics.Debugger.Launch();
            GenerateHelp(Assembly.GetExecutingAssembly(), outputPath, oneFile, null);
        }
        public static void GenerateHelp(string outputPath, bool oneFile, string cmdletName)
        {
            GenerateHelp(Assembly.GetExecutingAssembly(), outputPath, oneFile, cmdletName);
        }

        public static void GenerateHelp(string inputFile, string outputPath, bool oneFile)
        {
            GenerateHelp(Assembly.LoadFrom(inputFile), outputPath, oneFile, null);
        }
        public static void GenerateHelp(string inputFile, string outputPath, bool oneFile, string cmdletName)
        {
            GenerateHelp(Assembly.LoadFrom(inputFile), outputPath, oneFile, cmdletName);
        }

        public static void GenerateHelp(Assembly asm, string outputPath, bool oneFile)
        {
            GenerateHelp(asm, outputPath, oneFile, null);
        }
        public static void GenerateHelp(Assembly asm, string outputPath, bool oneFile, string cmdletName)
        {
            Console.WriteLine(string.Format("MamlGenerator.GenerateHelp(): asm={0}; outputPath={1}; oneFile={2}", asm.FullName, outputPath, oneFile));

            var attr = asm.GetCustomAttributes(typeof (AssemblyCopyrightAttribute), false);
            if (attr.Length == 1)
                _copyright = ((AssemblyCopyrightAttribute) attr[0]).Copyright;
            attr = asm.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
            if (attr.Length == 1)
                _copyright += ((AssemblyDescriptionAttribute)attr[0]).Description;


            StringBuilder sb = new StringBuilder();
            _writer = new XmlTextWriter(new StringWriter(sb));
            _writer.Formatting = Formatting.Indented;

            if (oneFile)
            {
                _writer.WriteStartElement("helpItems");
                _writer.WriteAttributeString("xmlns", "http://msh");
                _writer.WriteAttributeString("schema", "maml");
            }

            foreach (Type type in asm.GetExportedTypes())
            {
                CmdletAttribute ca = GetAttribute<CmdletAttribute>(type);
                if (ca != null)
                {
                    if (!string.IsNullOrEmpty(cmdletName) && cmdletName.ToLower() != string.Format("{0}-{1}", ca.VerbName, ca.NounName).ToLower()) continue;

                    Console.WriteLine(string.Format("MamlGenerator.GenerateHelp(): Found Cmdlet: {0}-{1}", ca.VerbName, ca.NounName));

                    if (!oneFile)
                    {
                        _writer.WriteStartElement("helpItems");
                        _writer.WriteAttributeString("xmlns", "http://msh");
                        _writer.WriteAttributeString("schema", "maml");
                    }


                    _writer.WriteStartElement("command", "command", "http://schemas.microsoft.com/maml/dev/command/2004/10");
                    _writer.WriteAttributeString("xmlns", "maml", null, "http://schemas.microsoft.com/maml/2004/1");
                    _writer.WriteAttributeString("xmlns", "dev", null, "http://schemas.microsoft.com/maml/dev/2004/10");
                    _writer.WriteAttributeString("xmlns", "gl", null, "http://schemas.falchionconsulting.com/maml/gl/2013/02");

                    _writer.WriteStartElement("command", "details", null);

                    _writer.WriteElementString("command", "name", null, string.Format("{0}-{1}", ca.VerbName, ca.NounName));

                    CmdletGroupAttribute group = GetAttribute<CmdletGroupAttribute>(type);
                    if (group != null && !string.IsNullOrEmpty(group.Group))
                        _writer.WriteElementString("gl", "group", null, group.Group);
                    else
                        _writer.WriteElementString("gl", "group", null, ca.NounName);

                    WriteDescription(type, true, false);

                    WriteCopyright();

                    _writer.WriteElementString("command", "verb", null, ca.VerbName);
                    _writer.WriteElementString("command", "noun", null, ca.NounName);

                    _writer.WriteElementString("dev", "version", null, asm.GetName().Version.ToString());

                    _writer.WriteEndElement(); //command:details

                    WriteDescription(type, false, true);

                    WriteSyntax(ca, type);
                    
                    _writer.WriteStartElement("command", "parameters", null);
                    
                    foreach (PropertyInfo pi in type.GetProperties())
                    {
                        List<ParameterAttribute> pas = GetAttribute<ParameterAttribute>(pi);
                        if (pas == null)
                            continue;

                        ParameterAttribute pa = null;
                        if (pas.Count == 1)
                            pa = pas[0];
                        else
                        {
                            // Determine the defualt property parameter set to use for details.
                            ParameterAttribute defaultPA = null;
                            foreach (ParameterAttribute temp in pas)
                            {
                                string defaultSet = ca.DefaultParameterSetName;
                                if (string.IsNullOrEmpty(ca.DefaultParameterSetName))
                                    defaultSet = string.Empty;

                                string set = temp.ParameterSetName;
                                if (string.IsNullOrEmpty(set) || set == ALL_PARAMETER_SETS_NAME)
                                {
                                    set = string.Empty;
                                    defaultPA = temp;
                                }
                                if (set.ToLower() == defaultSet.ToLower())
                                {
                                    pa = temp;
                                    defaultPA = temp;
                                    break;
                                }
                            }
                            if (pa == null && defaultPA != null)
                                pa = defaultPA;
                            if (pa == null)
                                pa = pas[0];
                        }
                        
                        _writer.WriteStartElement("command", "parameter", null);
                        _writer.WriteAttributeString("required", pa.Mandatory.ToString().ToLower());

                        bool supportsWildcard = GetAttribute<SupportsWildcardsAttribute>(pi) != null;
                        _writer.WriteAttributeString("globbing", supportsWildcard.ToString().ToLower());

                        if (!pa.ValueFromPipeline && !pa.ValueFromPipelineByPropertyName)
                            _writer.WriteAttributeString("pipelineInput", "false");
                        else if (pa.ValueFromPipeline && pa.ValueFromPipelineByPropertyName)
                            _writer.WriteAttributeString("pipelineInput", "true (ByValue, ByPropertyName)");
                        else if (!pa.ValueFromPipeline && pa.ValueFromPipelineByPropertyName)
                            _writer.WriteAttributeString("pipelineInput", "true (ByPropertyName)");
                        else if (pa.ValueFromPipeline && !pa.ValueFromPipelineByPropertyName)
                            _writer.WriteAttributeString("pipelineInput", "true (ByValue)");

                        if (pa.Position < 0)
                            _writer.WriteAttributeString("position", "named");
                        else
                            _writer.WriteAttributeString("position", (pa.Position + 1).ToString());

                        bool variableLength = pi.PropertyType.IsArray;
                        _writer.WriteAttributeString("variableLength", variableLength.ToString().ToLower());

                        _writer.WriteElementString("maml", "name", null, pi.Name);

                        if (pi.PropertyType.Name == "SPAssignmentCollection")
                            WriteSPAssignmentCollectionDescription();
                        else
                            WriteDescription(pa.HelpMessage, false);

                        _writer.WriteStartElement("command", "parameterValue", null);
                        _writer.WriteAttributeString("required", pa.Mandatory.ToString().ToLower());
                        _writer.WriteAttributeString("variableLength", variableLength.ToString().ToLower());
                        _writer.WriteValue(pi.PropertyType.Name);
                        _writer.WriteEndElement(); //command:parameterValue

                        WriteDevType(pi.PropertyType.Name, null);

                        _writer.WriteEndElement(); //command:parameter
                    }
                    _writer.WriteEndElement(); //command:parameters

                    //TODO: Find out what is supposed to go here
                    _writer.WriteStartElement("command", "inputTypes", null);
                    _writer.WriteStartElement("command", "inputType", null);
                    WriteDevType(null, null);
                    _writer.WriteEndElement(); //command:inputType
                    _writer.WriteEndElement(); //command:inputTypes

                    _writer.WriteStartElement("command", "returnValues", null);
                    _writer.WriteStartElement("command", "returnValue", null);
                    WriteDevType(null, null);
                    _writer.WriteEndElement(); //command:returnValue
                    _writer.WriteEndElement(); //command:returnValues

                    _writer.WriteElementString("command", "terminatingErrors", null, null);
                    _writer.WriteElementString("command", "nonTerminatingErrors", null, null);

                    _writer.WriteStartElement("maml", "alertSet", null);
                    _writer.WriteElementString("maml", "title", null, null);
                    _writer.WriteStartElement("maml", "alert", null);
                    WritePara(string.Format("For more information, type \"Get-Help {0}-{1} -detailed\". For technical information, type \"Get-Help {0}-{1} -full\".", 
                        ca.VerbName, ca.NounName));
                    _writer.WriteEndElement(); //maml:alert
                    _writer.WriteEndElement(); //maml:alertSet

                    WriteExamples(type);
                    WriteRelatedLinks(type);

                    _writer.WriteEndElement(); //command:command

                    if (!oneFile)
                    {
                        _writer.WriteEndElement(); //helpItems
                        _writer.Flush();
                        File.WriteAllText(Path.Combine(outputPath, string.Format("{0}.dll-help.xml", type.Name)), sb.ToString());
                        sb = new StringBuilder();
                        _writer = new XmlTextWriter(new StringWriter(sb));
                        _writer.Formatting = Formatting.Indented;
                    }
                }
            }

            if (oneFile)
            {
                _writer.WriteEndElement(); //helpItems
                _writer.Flush();
                File.WriteAllText(Path.Combine(outputPath, string.Format("{0}.dll-help.xml", asm.GetName().Name)), sb.ToString());
            }
        }

        const string ALL_PARAMETER_SETS_NAME = "__AllParameterSets";

        private static void WriteSyntax(CmdletAttribute ca, Type type)
        {
            Dictionary<string, List<PropertyInfo>> parameterSets = new Dictionary<string, List<PropertyInfo>>();

            List<PropertyInfo> allParameterSets = null;
            foreach (PropertyInfo pi in type.GetProperties())
            {
                List<ParameterAttribute> pas = GetAttribute<ParameterAttribute>(pi);
                if (pas == null)
                    continue;

                foreach (ParameterAttribute temp in pas)
                {
                    string set = temp.ParameterSetName + "";
                    List<PropertyInfo> piList = null;
                    if (!parameterSets.ContainsKey(set))
                    {
                        piList = new List<PropertyInfo>();
                        parameterSets.Add(set, piList);
                    }
                    else
                        piList = parameterSets[set];
                    parameterSets[set].Add(pi);
                }
            }
            if (parameterSets.Count == 0)
                return;

            if (parameterSets.ContainsKey(ALL_PARAMETER_SETS_NAME))
                allParameterSets = parameterSets[ALL_PARAMETER_SETS_NAME];

            if (parameterSets.Count > 1 && parameterSets.ContainsKey(ALL_PARAMETER_SETS_NAME))
                parameterSets.Remove(ALL_PARAMETER_SETS_NAME);
            if (parameterSets.Count == 1)
                allParameterSets = null;

            _writer.WriteStartElement("command", "syntax", null);
            foreach (string parameterSetName in parameterSets.Keys)
            {
                WriteSyntaxItem(ca, parameterSets, parameterSetName, allParameterSets);
            }
            _writer.WriteEndElement(); //command:syntax
        }

       
        private static void WriteSyntaxItem(CmdletAttribute ca, Dictionary<string, List<PropertyInfo>> parameterSets, string parameterSetName, List<PropertyInfo> allParameterSets)
        {
            _writer.WriteStartElement("command", "syntaxItem", null);
            _writer.WriteElementString("maml", "name", null, string.Format("{0}-{1}", ca.VerbName, ca.NounName));
            foreach (PropertyInfo pi in parameterSets[parameterSetName])
            {
                ParameterAttribute pa = GetParameterAttribute(pi, parameterSetName);
                if (pa == null)
                    continue;

                WriteParameter(pi, pa);
            }
            if (allParameterSets != null)
            {
                foreach (PropertyInfo pi in allParameterSets)
                {
                    List<ParameterAttribute> pas = GetAttribute<ParameterAttribute>(pi);
                    if (pas == null)
                        continue;
                    WriteParameter(pi, pas[0]);
                }
            }
            _writer.WriteEndElement(); //command:syntaxItem
        }
 
        private static ParameterAttribute GetParameterAttribute(PropertyInfo pi, string parameterSetName)
        {
            List<ParameterAttribute> pas = GetAttribute<ParameterAttribute>(pi);
            if (pas == null)
                return null;
            ParameterAttribute pa = null;
            foreach (ParameterAttribute temp in pas)
            {
                if (temp.ParameterSetName.ToLower() == parameterSetName.ToLower())
                {
                    pa = temp;
                    break;
                }
            }
            return pa;
        }

        private static void WriteParameter(PropertyInfo pi, ParameterAttribute pa)
        {
            _writer.WriteStartElement("command", "parameter", null);
            _writer.WriteAttributeString("required", pa.Mandatory.ToString().ToLower());
            //_writer.WriteAttributeString("parameterSetName", pa.ParameterSetName);
            if (pa.Position < 0)
                _writer.WriteAttributeString("position", "named");
            else
                _writer.WriteAttributeString("position", (pa.Position + 1).ToString());

            _writer.WriteElementString("maml", "name", null, pi.Name);
            _writer.WriteStartElement("command", "parameterValue", null);

            if (pi.DeclaringType == typeof(PSCmdlet))
                _writer.WriteAttributeString("required", "false");
            else
                _writer.WriteAttributeString("required", "true");

            if (pi.PropertyType.Name == "Nullable`1")
            {
                Type coreType = pi.PropertyType.GetGenericArguments()[0];
                if (coreType.IsEnum)
                    _writer.WriteValue(string.Join(" | ", Enum.GetNames(coreType)));
                else
                    _writer.WriteValue(coreType.Name);
            }
            else
            {
                if (pi.PropertyType.IsEnum)
                    _writer.WriteValue(string.Join(" | ", Enum.GetNames(pi.PropertyType)));
                else
                    _writer.WriteValue(pi.PropertyType.Name);
            }

            _writer.WriteEndElement(); //command:parameterValue
            _writer.WriteEndElement(); //command:parameter
        }

        private static void WriteDevType(string name, string description)
        {
            _writer.WriteStartElement("dev", "type", null);
            _writer.WriteElementString("maml", "name", null, name);
            _writer.WriteElementString("maml", "uri", null, null);
            WriteDescription(description, false);
            _writer.WriteEndElement(); //dev:type
        }

        private static void WriteSPAssignmentCollectionDescription()
        {
            WriteDescription("Manages objects for the purpose of proper disposal. Use of objects, such as SPWeb or SPSite, can use large amounts of memory and use of these objects in Windows PowerShell scripts requires proper memory management. Using the SPAssignment object, you can assign objects to a variable and dispose of the objects after they are needed to free up memory. When SPWeb, SPSite, or SPSiteAdministration objects are used, the objects are automatically disposed of if an assignment collection or the Global parameter is not used.\r\n\r\nWhen the Global parameter is used, all objects are contained in the global store. If objects are not immediately used, or disposed of by using the Stop-SPAssignment command, an out-of-memory scenario can occur.", false);
        }

        private static void WriteDescription(Type type, bool synopsis, bool addCopyright)
        {
            _writer.WriteStartElement("maml", "description", null);
            CmdletDescriptionAttribute da = GetAttribute<CmdletDescriptionAttribute>(type);
            string desc = string.Empty;
            if (synopsis)
            {
                if (da != null && !string.IsNullOrEmpty(da.Synopsis))
                {
                    desc = da.Synopsis;
                }
            }
            else
            {
                if (da != null && !string.IsNullOrEmpty(da.Description))
                {
                    desc = da.Description;
                }
            }

            WritePara(desc);
            if (addCopyright)
            {
                WritePara(null);
                WritePara(_copyright);
            }
            
            _writer.WriteEndElement(); //maml:description
        }

        private static void WriteDescription(string desc, bool addCopyright)
        {
            _writer.WriteStartElement("maml", "description", null);
            WritePara(desc);
            if (addCopyright)
            {
                WritePara(null);
                WritePara(_copyright);
            }
            _writer.WriteEndElement(); //maml:description
        }

        private static void WriteExamples(Type type)
        {
            object[] attrs = type.GetCustomAttributes(typeof(ExampleAttribute), true);
            if (attrs == null || attrs.Length == 0)
            {
                _writer.WriteElementString("command", "examples", null, null);
            }
            else
            {
                _writer.WriteStartElement("command", "examples", null);

                for (int i = 0; i < attrs.Length; i++)
                {
                    ExampleAttribute ex = (ExampleAttribute)attrs[i];
                    _writer.WriteStartElement("command", "example", null);
                    if (attrs.Length == 1)
                        _writer.WriteElementString("maml", "title", null, "------------------EXAMPLE------------------");
                    else
                        _writer.WriteElementString("maml", "title", null, string.Format("------------------EXAMPLE {0}-----------------------", i + 1));

                    _writer.WriteElementString("dev", "code", null, ex.Code);
                    _writer.WriteStartElement("dev", "remarks", null);
                    WritePara(ex.Remarks);
                    _writer.WriteEndElement(); //dev:remarks
                    _writer.WriteEndElement(); //command:example
                }
                _writer.WriteEndElement(); //command:examples
            }
        }

        private static void WriteRelatedLinks(Type type)
        {
            RelatedCmdletsAttribute attr = GetAttribute<RelatedCmdletsAttribute>(type);
            
            if (attr == null)
            {
                _writer.WriteElementString("maml", "relatedLinks", null, null);
            }
            else
            {
                _writer.WriteStartElement("maml", "relatedLinks", null);

                foreach (Type t in attr.RelatedCmdlets)
                {
                    CmdletAttribute ca = GetAttribute<CmdletAttribute>(t);
                    if (ca == null)
                        continue;

                    _writer.WriteStartElement("maml", "navigationLink", null);
                    _writer.WriteElementString("maml", "linkText", null, ca.VerbName + "-" + ca.NounName);
                    _writer.WriteElementString("maml", "uri", null, null);
                    _writer.WriteEndElement(); //maml:navigationLink
                }
                if (attr.ExternalCmdlets != null)
                {
                    foreach (string s in attr.ExternalCmdlets)
                    {
                        _writer.WriteStartElement("maml", "navigationLink", null);
                        _writer.WriteElementString("maml", "linkText", null, s);
                        _writer.WriteElementString("maml", "uri", null, null);
                        _writer.WriteEndElement(); //maml:navigationLink
                    }
                }
                _writer.WriteEndElement(); //maml:relatedLinks
            }
        }

        private static T GetAttribute<T>(Type type)
        {
            object[] attrs = type.GetCustomAttributes(typeof(T), true);
            if (attrs == null || attrs.Length == 0)
                return default(T);
            return (T)attrs[0];
        }
        private static List<T> GetAttribute<T>(PropertyInfo pi)
        {
            object[] attrs = pi.GetCustomAttributes(typeof(T), true);
            List<T> attributes = new List<T>();
            if (attrs == null || attrs.Length == 0)
                return null;

            foreach (T t in attrs)
            {
                attributes.Add(t);
            }
            return attributes;
        }

        private static void WriteCopyright()
        {
            _writer.WriteStartElement("maml", "copyright", null);
            WritePara(_copyright);
            _writer.WriteEndElement(); //maml:copyright
        }

        private static void WritePara(string para)
        {
            if (string.IsNullOrEmpty(para))
            {
                _writer.WriteElementString("maml", "para", null, null);
                return;
            }
            string[] paragraphs = para.Split(new[] {"\r\n"}, StringSplitOptions.None);
            foreach (string p in paragraphs)
                _writer.WriteElementString("maml", "para", null, p);
        }
    }
}
