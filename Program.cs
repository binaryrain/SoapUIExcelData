using Microsoft.CSharp;
using Microsoft.Office.Interop.Excel;
using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace XML2SoapXML
{

    public delegate void delNotifyFileCompletion(string fileName);
    public class Program
    {
        private Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        private Dictionary<string, string> sheets = new Dictionary<string, string>();
        private Dictionary<string, int> columnTrack = new Dictionary<string, int>();
        private Assembly asm;
        private static string wsdlUrl = "";
        private static string samplexmlname = "";
        private static string methodNameRequestFolder = "";
        public event delNotifyFileCompletion OnNotifyFileCompleted;
        private List<string> doneSheets = new List<string>();
        private Dictionary<string, object> asmTypes = new Dictionary<string, object>();
        //private Dictionary<string, string> arrayPropsId = new Dictionary<string, string>();
        private Hashtable localCurrentTypesCache = new Hashtable();
        XmlDocument xmlSampleDoc = null;

        static void Main(string[] args)
        {
            #region Commented Code
            //ServiceReference1.First_SumService_WSD_PortTypeClient c = new ServiceReference1.First_SumService_WSD_PortTypeClient();
            //string gg = c.First_SumService("2", "4");
            //var t = gg;
            //XmlReader reader = XmlReader.Create(@"C:\Users\ja13\Desktop\xsd\1301.xml");
            //string currentFileName = "1301";

            //XmlSchemaSet schemaSet = new XmlSchemaSet();
            //XmlSchemaInference schema = new XmlSchemaInference();

            //schemaSet = schema.InferSchema(reader);
            //int i = 1;
            //foreach (System.Xml.Schema.XmlSchema s in schemaSet.Schemas())
            //{
            //    StreamWriter writer = new StreamWriter(@"C:\Users\ja13\Desktop\xsd\"+currentFileName+i+".xsd");
            //    s.Write(writer);
            //    writer.Flush();
            //    writer.Close();
            //    i++;
            //}
            #endregion

            try
            {

                wsdlUrl = args[0];
                Program p = new Program(wsdlUrl);
                if (args.Count() == 2)
                {
                    if (args[1].EndsWith(".xml"))
                    {
                        p.createExcelProcess(args[1]);
                    }
                }
                else
                {
                    samplexmlname = args[1];
                    methodNameRequestFolder = args[2];
                    p.ReadExcel();
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(Environment.CurrentDirectory + "\\" + "dffd2.txt", ex.Message + "------------" + ex.StackTrace);
                // throw;
                // StreamWriter wr = File.CreateText(Environment.CurrentDirectory + "\\" + "dffd2.txt");
                // wr.WriteLine(ex.Message + "---------------" + ex.StackTrace);
                // wr.Close();
            }

        }

        public Program(string wsdl)
        {
            wsdlUrl = wsdl;
        }

        public bool createExcelProcess(string xml)
        {
            bool resBool = false;
            try
            {
                Process procXSD = Process.Start(Environment.CurrentDirectory + "\\" + "xsd.exe", "\"" + xml + "\"");

                procXSD.WaitForExit();
                //   File.AppendAllText(Environment.CurrentDirectory + "\\" + "dffd2.txt", ": endte2");
                XDocument doc;
                try
                {
                    doc = XDocument.Load(wsdlUrl);
                }
                catch
                {
                    Console.WriteLine("Service not accessible");
                    throw;
                }
                //File.AppendAllText(Environment.CurrentDirectory + "\\" + "dffd.txt", xml);
                List<XElement> ElementsMultiple = new List<XElement>();


                XElement ele = doc.Descendants().SingleOrDefault(p => p.Name.LocalName == "types");
                List<XElement> elements = ele.Descendants().Where(t => t.Name.LocalName == "element").ToList();
                foreach (XElement e in elements)
                {
                    if (e.Attributes().Count(t => t.Name == "maxOccurs") > 0)
                    {
                        if (e.Attributes().FirstOrDefault(t => t.Name == "maxOccurs").Value == "unbounded")
                        {
                            ElementsMultiple.Add(e);
                        }
                    }
                }

                string[] Xsds = Directory.GetFiles(Environment.CurrentDirectory, "*.xsd");
                foreach (string f in Xsds)
                {
                    XDocument XsdDoc = XDocument.Load(f);
                    List<XElement> XsdElements = XsdDoc.Descendants().Where(p => p.Name.LocalName == "element").ToList();

                    {
                        foreach (XElement e in XsdElements)
                        {
                            if (e.Attributes().Count(t => t.Name == "maxOccurs") > 0)
                            {
                                if (e.Attributes().FirstOrDefault(t => t.Name == "maxOccurs").Value == "unbounded")
                                {
                                    bool fnd = false;
                                    if (ElementsMultiple.Count(t => t.Attribute(XName.Get("name")).Value == e.Attribute(XName.Get("name")).Value) > 0)
                                    {
                                        fnd = true;
                                    }
                                    else if (e.Descendants().Count(p => p.Name.LocalName == "complexType") > 0)
                                    {
                                        XElement complexTypeElement = e.Descendants().FirstOrDefault(p => p.Name.LocalName == "element");
                                        if (complexTypeElement != null)
                                        {
                                            if (complexTypeElement.Attributes().Count(y => y.Name == "ref") == 1)
                                            {
                                                //if (ElementsMultiple.Count(t => t.Attribute(XName.Get("name")).Value == complexTypeElement.Attribute(XName.Get("ref")).Value) > 0)
                                                //{
                                                foreach (XElement elet in ElementsMultiple)
                                                {
                                                    if (complexTypeElement.Attribute(XName.Get("ref")).Value.Contains(elet.Attribute(XName.Get("name")).Value))
                                                    {
                                                        fnd = true;
                                                        break;
                                                    }
                                                }
                                                // }
                                            }
                                        }
                                    }
                                    if (!fnd)
                                    {
                                        foreach (var attribute in e.Attributes())
                                        {
                                            if (attribute.Name == "maxOccurs")
                                                attribute.Remove();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    File.Delete(f);
                    XsdDoc.Save(f);
                }

                Xsds = Directory.GetFiles(Environment.CurrentDirectory, "*.xsd");
                Xsds = Xsds.OrderBy(t => t).ToArray();
                string xsdToClass = Environment.CurrentDirectory + "\\" + "xsd.exe";
                string parameter = "/c ";
                foreach (string i in Xsds)
                    parameter = parameter + "\"" + i + "\"" + " ";
                Process proc = Process.Start(xsdToClass, parameter);
                proc.WaitForExit();

                string[] classes = Directory.GetFiles(Environment.CurrentDirectory, "*.cs");
                string classFileCode = File.ReadAllText(classes[0]);
                string source = classes[0];

                var provider = new CSharpCodeProvider();
                var options = new CompilerParameters
                {
                    GenerateExecutable = false
                };
                options.ReferencedAssemblies.Add("System.dll");
                options.ReferencedAssemblies.Add("System.Xml.dll");


                var res = provider.CompileAssemblyFromFile(options, new[] { source });


                asm = res.CompiledAssembly;

                sheets.Add("Root", "");
                //p.currentsheet = "Root";

                string rootname = "";

                xlApp.Visible = false;

                Workbook wBook = xlApp.Workbooks.Add(Type.Missing);

                Worksheet rootsheet = (Microsoft.Office.Interop.Excel.Worksheet)wBook.ActiveSheet;
                rootsheet.Name = "Root";

                var newDataset = asm.GetTypes().FirstOrDefault(m => m.GetProperties().Any(o => o.PropertyType == (new object[] { }).GetType()));
                if (newDataset != null)
                {
                    rootname = newDataset.Name;
                }
                else
                {
                    foreach (Type t in asm.GetTypes())
                    {
                        IEnumerable<CustomAttributeData> c = t.CustomAttributes;

                        if (c.Count(g => g.AttributeType.Name == "XmlRootAttribute") == 1)
                        {
                            // find one class that is not a property in other class.. it would be Root class
                            if (asm.GetTypes().Any(m => m.GetProperties().Any(h => (h.PropertyType.IsArray && h.PropertyType.GetElementType() == t) || (!h.PropertyType.IsArray && h.PropertyType == t))))
                            {

                            }
                            else
                            {
                                rootname = t.Name;
                                break;
                            }
                        }
                    }
                }

                CreateExcelSheet(asm.GetType(rootname), "Root", "Root", "");
                xlApp.ActiveWorkbook.SaveAs(Environment.CurrentDirectory + "\\" + "mysheet2.xlsx");

                xlApp.ActiveWorkbook.Close();

                resBool = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return resBool;

        }

        public void CreateExcelSheet(Type root, string currentsheet, string previousSheet, string caption)
        {

            PropertyInfo[] propreties = root.GetProperties();

            foreach (PropertyInfo prop in propreties)
            {
                if (prop.PropertyType == typeof(object) || prop.PropertyType.GetElementType() == typeof(object))
                {
                    if (sheets.ContainsKey(prop.Name))
                    {
                        if (!sheets[currentsheet].Contains(prop.Name + "_id"))
                        {
                            if (columnTrack.Keys.Contains(currentsheet))
                            {
                                columnTrack[currentsheet] = columnTrack[currentsheet] + 1;
                            }
                            else
                            {
                                columnTrack.Add(currentsheet, 1);
                            }
                            ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]] = prop.Name + "_id";
                            if (currentsheet == "Root")
                            {
                                ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, 2] = "Test Case Name";
                            }
                        }
                        continue;
                    }
                    Worksheet sheet = (Worksheet)xlApp.ActiveWorkbook.Worksheets.Add();
                    sheet.Name = prop.Name;

                    if (columnTrack.Keys.Contains(currentsheet))
                    {
                        columnTrack[currentsheet] = columnTrack[currentsheet] + 1;
                    }
                    else
                    {
                        columnTrack.Add(currentsheet, 1);
                    }

                    ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]] = prop.Name + "_id";
                    ((Range)((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]]).AddComment(caption);//root.Name

                    if (currentsheet == "Root")
                    {
                        ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, 2] = "Test Case Name";
                    }

                    sheets[currentsheet] = sheets[currentsheet] + "," + prop.Name + "_id";
                    sheets.Add(prop.Name, prop.Name + "_id");
                    previousSheet = currentsheet;
                    currentsheet = prop.Name;

                    if (columnTrack.Keys.Contains(currentsheet))
                    {
                        columnTrack[currentsheet] = columnTrack[currentsheet] + 1;
                    }
                    else
                    {
                        columnTrack.Add(currentsheet, 1);
                    }

                    ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]] = prop.Name + "_id";
                    // caption = caption + "," + prop.PropertyType.GetElementType().Name;
                    ((Range)((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]]).AddComment(prop.PropertyType.GetElementType().Name);//

                    IEnumerable<CustomAttributeData> atts = prop.CustomAttributes;
                    foreach (CustomAttributeData att in atts)
                    {
                        if (att.AttributeType.Name == "XmlElementAttribute" && att.Constructor.GetParameters().Count() == 2)
                        {
                            CustomAttributeTypedArgument p = att.ConstructorArguments[1];
                            caption = caption + "," + p.Value.ToString();
                            CreateExcelSheet(asm.GetType(p.Value.ToString()), currentsheet, previousSheet, caption);
                            //  currentsheet = previousSheet;
                            caption.Replace("," + p.Value.ToString(), "");
                        }
                    }
                }

                else if (prop.PropertyType == typeof(string) || prop.PropertyType == typeof(int) || prop.PropertyType == typeof(bool) || prop.PropertyType == typeof(decimal)
                    || prop.PropertyType == typeof(double) || prop.PropertyType == typeof(float) || prop.PropertyType == typeof(long) || prop.PropertyType == typeof(DateTime)
                    || prop.PropertyType == typeof(char))  //TODO Add more here\
                {

                    if (columnTrack.Keys.Contains(currentsheet))
                    {
                        columnTrack[currentsheet] = columnTrack[currentsheet] + 1;
                    }
                    else
                    {
                        columnTrack.Add(currentsheet, 1);
                    }
                    ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]] = prop.Name;
                    if (!caption.Contains(root.Name))
                    {
                        caption = caption + "," + root.Name;
                    }
                    ((Range)((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]]).AddComment(caption);//root.Name
                    if (currentsheet == "Root")
                    {
                        ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, 2] = "Test Case Name";
                    }
                    sheets[currentsheet] = sheets[currentsheet] + "," + prop.Name + "(" + root.Name + ")";
                }
                else if (!prop.PropertyType.IsArray)
                {
                    CreateExcelSheet(prop.PropertyType, currentsheet, previousSheet, caption + "," + prop.PropertyType.Name);
                    caption = caption.Replace("," + prop.PropertyType.Name, "");
                }
                else
                {
                    if (sheets.ContainsKey(prop.Name))
                    {
                        if (!sheets[currentsheet].Contains(prop.Name + "_id"))
                        {
                            if (columnTrack.Keys.Contains(currentsheet))
                            {
                                columnTrack[currentsheet] = columnTrack[currentsheet] + 1;
                            }
                            else
                            {
                                columnTrack.Add(currentsheet, 1);
                            }
                            ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]] = prop.Name + "_id";
                        }
                        if (currentsheet == "Root")
                        {
                            ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, 2] = "Test Case Name";
                        }
                        continue;
                    }
                    Worksheet sheet = (Worksheet)xlApp.ActiveWorkbook.Worksheets.Add();
                    sheet.Name = prop.Name;

                    if (columnTrack.Keys.Contains(currentsheet))
                    {
                        columnTrack[currentsheet] = columnTrack[currentsheet] + 1;
                    }
                    else
                    {
                        columnTrack.Add(currentsheet, 1);
                    }

                    ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]] = prop.Name + "_id";
                    // caption = caption + "," + root.Name;
                    ((Range)((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]]).AddComment(caption);//root.Name

                    if (currentsheet == "Root")
                    {
                        ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, 2] = "Test Case Name";
                    }

                    sheets[currentsheet] = sheets[currentsheet] + "," + prop.Name + "_id";
                    sheets.Add(prop.Name, prop.Name + "_id");
                    previousSheet = currentsheet;
                    currentsheet = prop.Name;

                    if (columnTrack.Keys.Contains(currentsheet))
                    {
                        columnTrack[currentsheet] = columnTrack[currentsheet] + 1;
                    }
                    else
                    {
                        columnTrack.Add(currentsheet, 1);
                    }

                    ((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]] = prop.Name + "_id";
                    // caption = caption + "," + prop.PropertyType.GetElementType().Name;
                    ((Range)((Worksheet)xlApp.ActiveWorkbook.Worksheets[currentsheet]).Cells[1, columnTrack[currentsheet]]).AddComment(prop.PropertyType.GetElementType().Name);//

                    CreateExcelSheet(prop.PropertyType.GetElementType(), currentsheet, previousSheet, prop.PropertyType.GetElementType().Name);

                    currentsheet = previousSheet;

                }
            }

        }

        List<System.Data.DataTable> dts = new List<System.Data.DataTable>();

        public bool ReadExcel()
        {
            bool result = false;
            Workbook wBook = null;
            try
            {
                string[] excelFiles = Directory.GetFiles(Environment.CurrentDirectory, "*.xlsx");
                if (excelFiles.Count() != 1)
                {
                    return false;
                }

                Application app = new Application();
                app.Visible = false;

                wBook = app.Workbooks.Open(excelFiles[0]);

                foreach (Worksheet wSheet in wBook.Worksheets)
                {
                    WorksheetSheetToDatatable(wSheet);
                }

                string[] classes = Directory.GetFiles(Environment.CurrentDirectory, "*.cs");

                var provider = new CSharpCodeProvider();
                var options = new CompilerParameters
                {
                    GenerateExecutable = false
                };
                options.ReferencedAssemblies.Add("System.dll");
                options.ReferencedAssemblies.Add("System.Xml.dll");
                string source = classes[0];

                var res = provider.CompileAssemblyFromFile(options, new[] { source });
                asm = res.CompiledAssembly;

                NavigateSheetData("Root", "", ref localCurrentTypesCache);
                result = true;
            }
            catch
            {
                throw;
            }
            finally
            {
                wBook.Close();
            }
            return result;
        }

        private void WorksheetSheetToDatatable(Worksheet sheet)
        {
            try
            {
                Range xlRange = sheet.UsedRange;

                System.Data.DataTable dt = new System.Data.DataTable(sheet.Name);


                for (int column = 1; column <= sheet.UsedRange.Columns.Count; ++column)
                {
                    Range r = sheet.Cells[1, column];
                    if (r.Value == null)
                        continue;
                    string val = (string)r.Value;
                    string com = string.Empty;
                    Comment p = r.Comment;
                    if (p != null && p.Text() != null)
                    {
                        com = p.Text();
                    }
                    System.Data.DataColumn c = new System.Data.DataColumn();
                    c.Caption = com;
                    if (com.Contains(','))
                        com = com.Substring(com.LastIndexOf(',') + 1);
                    c.ColumnName = val + "$$" + com;
                    dt.Columns.Add(c);
                }

                for (int row = 2; row <= sheet.UsedRange.Rows.Count; ++row)
                {
                    System.Data.DataRow _row = dt.NewRow();
                    for (int column = 1; column <= sheet.UsedRange.Columns.Count; ++column)
                    {
                        Range r = sheet.Cells[row, column];
                        string val = Convert.ToString(r.Value);
                        if (r.Value == null)
                            continue;
                        _row[column - 1] = val;
                    }
                    dt.Rows.Add(_row);
                }

                dts.Add(dt);
            }
            catch
            {
                throw;
            }
        }

        static string curFileName = "";
        //static string uniueidStat = Guid.NewGuid().ToString("D");
        private void NavigateSheetData(string TableName, string CurrentId, ref Hashtable localCurrentTypesCache)
        {
            // sheet is already been read. So do not let recursion repeat sheet.
            if (!doneSheets.Contains(TableName + CurrentId))
            {
                doneSheets.Add(TableName + CurrentId);
            }
            else
            {
                return;
            }
            System.Data.DataTable dt = dts.FirstOrDefault(t => t.TableName == TableName);
            if (dt != null)
            {
                System.Data.DataRowCollection rows = dt.Rows;
                foreach (System.Data.DataRow row in rows)
                {
                    for (int i = 0; i < row.Table.Columns.Count; i++)
                    {
                        if (TableName == "Root")    // read test case name
                        {
                            if (dt.Columns.Count == 2)
                            {
                                if (string.IsNullOrEmpty(row[1].ToString()))
                                {
                                    int rand = new Random().Next();
                                    curFileName = rand + ".xml";
                                }
                                else
                                {
                                    curFileName = row[1].ToString() + ".xml";
                                }
                            }
                            else
                            {
                                int rand = new Random().Next();
                                curFileName = rand + ".xml";
                            }
                        }
                        if (row.Table.Columns[i].ColumnName.Split(new string[] { "$$" }, StringSplitOptions.None)[0].Contains("_id"))
                        {
                            if (TableName + "_id" != row.Table.Columns[i].ColumnName.Split(new string[] { "$$" }, StringSplitOptions.None)[0]) //need to read column's sheet and come back to parent sheet later
                            {
                                if (row[0].ToString() == CurrentId || CurrentId == "")
                                    NavigateSheetData(row.Table.Columns[i].ColumnName.Split(new string[] { "$$" }, StringSplitOptions.None)[0].Replace("_id", ""), row[i].ToString(), ref localCurrentTypesCache);
                            }
                            else if (TableName + "_id" == row.Table.Columns[i].ColumnName.Split(new string[] { "$$" }, StringSplitOptions.None)[0])
                            {
                                if (row[0].ToString() != CurrentId || CurrentId == "") // no need to read other xml request data
                                {
                                    continue;
                                }
                                string caption = row.Table.Columns[i].Caption;
                                Type lst = typeof(List<>); // it will always be list as sheet name is equal to column name
                                Type constructedListType = null;
                                if (caption.ToLower() == "object")
                                {
                                    constructedListType = lst.MakeGenericType(typeof(object));
                                }
                                else
                                {
                                    constructedListType = lst.MakeGenericType(asm.GetType(caption));
                                }
                                var instance = Activator.CreateInstance(constructedListType);
                                if (asmTypes.Count(t => t.Key == instance.GetType().FullName + CurrentId) == 0)
                                {
                                    asmTypes = asmTypes.OrderBy(t => Convert.ToInt32(t.Key.Substring(t.Key.LastIndexOf(']') + 1))).Select(t => t).ToDictionary(t => t.Key, t => t.Value);
                                    if (asmTypes.Count(t => t.Value.GetType() == instance.GetType()) > 0)
                                    {
                                        var lastTypeAdded = asmTypes.Last(t => t.Value.GetType() == instance.GetType());

                                        int lastIndexOfBracket = lastTypeAdded.Key.LastIndexOf(']');
                                        string lastTypeNumber = lastTypeAdded.Key.Substring(lastIndexOfBracket + 1);
                                        while (Convert.ToInt32(CurrentId) - Convert.ToInt32(lastTypeNumber) != 1)
                                        {
                                            string idToadd = (Convert.ToInt32(lastTypeNumber) + 1).ToString();
                                            var instanceEmpty = Activator.CreateInstance(constructedListType);
                                            lastTypeNumber = idToadd;
                                        }
                                        asmTypes.Add(instance.GetType().FullName + CurrentId, instance);
                                    }
                                    else
                                    {
                                        asmTypes.Add(instance.GetType().FullName + CurrentId, instance);
                                    }
                                }

                                System.Data.DataRow[] drs = dt.Select(row.Table.Columns[i].ColumnName + "=" + "'" + CurrentId + "'");
                                foreach (System.Data.DataRow r in drs)
                                {
                                    string uniqueId = Guid.NewGuid().ToString("D");
                                    object objType = null;
                                    if (caption.ToLower() == "object")
                                    {
                                        objType = Activator.CreateInstance(typeof(object));
                                    }
                                    else
                                    {
                                        objType = Activator.CreateInstance(asm.GetType(caption));
                                    }
                                    instance.GetType().GetMethod("Add").Invoke(instance, new[] { objType });
                                    System.Data.DataColumnCollection cls = r.Table.Columns;

                                    object lastType = null;

                                    for (int colCount = 1; colCount < cls.Count; colCount++)
                                    {
                                        var c = cls[colCount];
                                        if (c.ColumnName.Split(new string[] { "$$" }, StringSplitOptions.None)[0].Contains("_id"))
                                        {
                                            NavigateSheetData(c.ColumnName.Split(new string[] { "$$" }, StringSplitOptions.None)[0].Replace("_id", ""), r[colCount].ToString(), ref localCurrentTypesCache);
                                        }
                                        else
                                        {
                                            if (c.Caption == objType.GetType().Name)
                                            {
                                                var obj = (Array)instance.GetType().GetMethod("ToArray").Invoke(instance, new object[] { });
                                                var lastObj = obj.GetValue(obj.Length - 1);//GetType().GetMethod("GetValue").Invoke(obj, new object[] { 1 });
                                                lastObj.GetType().GetProperty(c.ColumnName.Split(new string[] { "$$" }, StringSplitOptions.None)[0]).SetValue(lastObj, r[colCount].ToString());
                                                if (!localCurrentTypesCache.HashTableKeyExist(lastObj, CurrentId + uniqueId))
                                                {
                                                    localCurrentTypesCache.Add(lastObj, CurrentId + uniqueId);
                                                }
                                            }
                                            else
                                            {
                                                var obj = (Array)instance.GetType().GetMethod("ToArray").Invoke(instance, new object[] { });
                                                var lastObj = obj.GetValue(obj.Length - 1);//GetType().GetMethod("GetValue").Invoke(obj, new object[] { 1 });
                                                lastType = lastObj;

                                                var props = lastType.GetType().GetProperties();
                                                foreach (var prop in props)
                                                {
                                                    if (localCurrentTypesCache.HashTableKeyExist(prop, CurrentId + uniqueId) && prop.GetValue(lastType) == null)
                                                    {
                                                        var obje = localCurrentTypesCache.HashTableGetKey(prop, CurrentId + uniqueId);
                                                        prop.SetValue(lastType, obje);
                                                    }
                                                }
                                                //if (localCurrentTypesCache.HashTableKeyExist(lastObj))
                                                //{
                                                //    objType = localCurrentTypesCache.HashTableGetKey(lastObj);
                                                //    instance.GetType().GetMethod("Add").Invoke(instance, new[] { objType });

                                                //}

                                                string[] classesNames = c.Caption.Split(new char[] { ',' }).Where(t => !String.IsNullOrEmpty(t)).Reverse().ToArray();

                                                foreach (string classname in classesNames)
                                                {
                                                    var currentType = Activator.CreateInstance(asm.GetType(classname));

                                                    if (localCurrentTypesCache.HashTableKeyExist(currentType, CurrentId + uniqueId))
                                                    {
                                                        currentType = localCurrentTypesCache.HashTableGetKey(currentType, CurrentId + uniqueId);
                                                    }

                                                    if (currentType.GetType().GetProperties().Any(t => t.Name == c.ColumnName.Split(new string[] { "$$" }, StringSplitOptions.None)[0]))
                                                    {
                                                        currentType.GetType().GetProperty(c.ColumnName.Split(new string[] { "$$" }, StringSplitOptions.None)[0]).SetValue(currentType, r[colCount].ToString());
                                                        lastType = currentType;
                                                        if (!localCurrentTypesCache.HashTableKeyExist(currentType, CurrentId + uniqueId))
                                                        {
                                                            localCurrentTypesCache.Add(currentType, CurrentId + uniqueId);
                                                        }
                                                        else
                                                        {
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (currentType.GetType().GetProperties().Any(t => t.PropertyType == lastType.GetType()))
                                                        {
                                                            PropertyInfo pInfo = currentType.GetType().GetProperties().First(t => t.PropertyType == lastType.GetType());
                                                            currentType.GetType().GetProperty(pInfo.Name).SetValue(currentType, lastType);
                                                            lastType = currentType;
                                                            if (currentType.GetType() == lastObj.GetType())
                                                            {
                                                                lastObj = currentType;
                                                            }
                                                            if (!localCurrentTypesCache.HashTableKeyExist(currentType, CurrentId + uniqueId))
                                                            {
                                                                localCurrentTypesCache.Add(currentType, CurrentId + uniqueId);
                                                            }
                                                            else
                                                            {
                                                                break;
                                                            }
                                                        }
                                                    }

                                                }
                                            }
                                        }
                                    }
                                }

                            }
                            else if (TableName == "Root")
                            {

                            }
                        }
                    } // for column end
                    if (TableName == "Root")
                    {
                        var newDataset = asm.GetTypes().FirstOrDefault(m => m.GetProperties().Any(o => o.PropertyType == (new object[] { }).GetType()));
                        if (newDataset != null)
                        {
                            CreateSoapXml(newDataset);
                        }
                        else
                        {
                            foreach (Type t in asm.GetTypes())
                            {
                                IEnumerable<CustomAttributeData> c = t.CustomAttributes;
                                if (c.Count(g => g.AttributeType.Name == "XmlRootAttribute") == 1)
                                {
                                    if (asm.GetTypes().Any(m => m.GetProperties().Any(h => (h.PropertyType.IsArray && h.PropertyType.GetElementType() == t) || (!h.PropertyType.IsArray && h.PropertyType == t))))
                                    {

                                    }
                                    else
                                    {
                                        CreateSoapXml(t);
                                        break;
                                    }
                                }
                            }
                        }

                        doneSheets.Clear();
                        localCurrentTypesCache = new Hashtable();
                        asmTypes.Clear();
                    }
                }  // foreach row ends


            }
        }

        private void CreateSoapXml(Type t)
        {
            asmTypes = asmTypes.OrderBy(m => Convert.ToInt32(m.Key.Substring(m.Key.LastIndexOf(']') + 1))).Select(m => m).ToDictionary(m => m.Key, m => m.Value); // need to sort this order by logic
            var yu = asm.CreateInstance(t.Name);
            PrepareObjectToSerialize(yu);

            XmlSerializer x = new XmlSerializer(yu.GetType());

            string targetNamespace = "";
            XDocument doc = XDocument.Load(wsdlUrl);
            List<XElement> ElementsMultiple = new List<XElement>();

            XElement eleTargetNS = doc.Descendants().SingleOrDefault(p => p.Name.LocalName == "definitions");
            targetNamespace = eleTargetNS.Attributes().FirstOrDefault(f => f.Name == "targetNamespace").Value;

            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("par", targetNamespace);
            ns.Add("soapenv", "http://schemas.xmlsoap.org/soap/envelope/");

            // int rand = new Random().Next();
            if (!Directory.Exists(Environment.CurrentDirectory + "\\" + methodNameRequestFolder))
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + "\\" + methodNameRequestFolder);
            }

            TextWriter writer = new StreamWriter(String.Format(Environment.CurrentDirectory + "\\" + methodNameRequestFolder + "\\" + curFileName));

            x.Serialize(writer, yu, ns);
            writer.Close();
            //curFileName = "Request" + rand + ".xml";
            if (File.Exists(String.Format(Environment.CurrentDirectory + "\\" + methodNameRequestFolder + "\\" + curFileName)))
            {
                XDocument xmlDoc = XDocument.Load(String.Format(Environment.CurrentDirectory + "\\" + methodNameRequestFolder + "\\" + curFileName));
                if (xmlDoc.Root.Name.LocalName == "NewDataSet")
                {
                    List<XAttribute> atts = xmlDoc.Root.Attributes().ToList();
                    if (xmlDoc.Root.Descendants().Count(g => g.Name.LocalName == "Envelope") > 0)
                    {
                        XElement enve = xmlDoc.Root.Descendants().First(g => g.Name.LocalName == "Envelope");
                        foreach (var att in atts)
                        {
                            enve.Add(att);
                        }
                        xmlDoc.Root.Remove();
                        xmlDoc.AddFirst(enve);
                    }
                    File.Delete(String.Format(Environment.CurrentDirectory + "\\" + methodNameRequestFolder + "\\" + curFileName));

                    xmlDoc.Save(String.Format(Environment.CurrentDirectory + "\\" + methodNameRequestFolder + "\\" + curFileName));
                }
                XmlReaderSettings readerSettings = new XmlReaderSettings();
                readerSettings.IgnoreComments = true;
                using (XmlReader reader = XmlReader.Create(samplexmlname, readerSettings))
                {
                    xmlSampleDoc = new XmlDocument();
                    xmlSampleDoc.Load(reader);
                }
                XmlDocument originalDoc = new XmlDocument();
                originalDoc.Load(String.Format(Environment.CurrentDirectory + "\\" + methodNameRequestFolder + "\\" + curFileName));
                XmlElement fRoot = originalDoc.DocumentElement;
                XmlNode root = fRoot.CloneNode(true);

                doOrdering(root);
                using (XmlWriter wri = XmlWriter.Create(String.Format(Environment.CurrentDirectory + "\\" + methodNameRequestFolder + "\\" + curFileName)))
                {
                    //originalDoc.RemoveAll();
                    root.WriteTo(wri);
                }

                // TODO : order elements as per sample xml

            }

            if (OnNotifyFileCompleted != null)
            {
                OnNotifyFileCompleted(curFileName);

            }
        }

        private void PrepareObjectToSerialize(object obj)
        {
            if (obj.GetType().IsArray)
            {
                var multipleProps = (IList)obj;
                foreach (var item in multipleProps)
                {
                    foreach (var prop in item.GetType().GetProperties())
                    {
                        if (prop.GetValue(item) == null)
                        {
                            if (prop.PropertyType.IsArray)
                            {
                                Type lst = typeof(List<>);
                                var constructedListType = lst.MakeGenericType(prop.PropertyType.GetElementType());

                                var instance = Activator.CreateInstance(constructedListType);

                                if (asmTypes.Count(t => t.Value.GetType() == instance.GetType()) > 0)
                                {
                                    // asmTypes = asmTypes.OrderBy(t => t.Key).ThenBy(t => Convert.ToInt32(t.Key.Substring(t.Key.LastIndexOf(']') + 1))).Select(t => t).ToDictionary(t => t.Key, t => t.Value);
                                    var arrObj = asmTypes.First(t => t.Value.GetType() == instance.GetType());
                                    var val = arrObj.Value;
                                    // if (asmTypes.Count(t => t.Value.GetType() == instance.GetType()) > 1)
                                    asmTypes.Remove(arrObj.Key);
                                    var y = val.GetType().GetMethod("ToArray").Invoke(val, new object[] { });
                                    prop.SetValue(item, y);
                                }
                                else
                                { 
                                
                                }
                            }
                            else
                            {

                                var objType = Activator.CreateInstance(prop.PropertyType);

                                //if (localCurrentTypesCache.HashTableKeyExist(objType))
                                //{
                                //    objType = localCurrentTypesCache.HashTableGetKey(objType);
                                //}
                                prop.SetValue(item, objType);


                            }
                            if (prop.GetValue(item) != null && prop.PropertyType != typeof(string))
                                PrepareObjectToSerialize(prop.GetValue(item));
                        }
                        else
                        {
                            if (prop.PropertyType != typeof(string) && prop.PropertyType != typeof(int) && prop.PropertyType != typeof(bool) && prop.PropertyType != typeof(decimal)
                                && prop.PropertyType != typeof(double) && prop.PropertyType != typeof(float) && prop.PropertyType != typeof(long) && prop.PropertyType != typeof(DateTime)
                                && prop.PropertyType != typeof(char)
                                && !prop.PropertyType.IsArray)
                            {
                                PrepareObjectToSerialize(prop.GetValue(item));
                            }
                        }
                    }
                }
            }
            else
            {
                foreach (var prop in obj.GetType().GetProperties())
                {
                    if (prop.GetValue(obj) == null)
                    {
                        if (prop.PropertyType.IsArray && prop.PropertyType == (new object[] { }).GetType())
                        {
                            Type lst = typeof(List<>);
                            var constructedListType = lst.MakeGenericType(typeof(object));

                            var instance = Activator.CreateInstance(constructedListType);

                            IEnumerable<CustomAttributeData> atts = prop.CustomAttributes;
                            foreach (CustomAttributeData att in atts)
                            {
                                if (att.AttributeType.Name == "XmlElementAttribute" && att.Constructor.GetParameters().Count() == 2)
                                {
                                    CustomAttributeTypedArgument p = att.ConstructorArguments[1];

                                    var objType = Activator.CreateInstance(asm.GetType(p.Value.ToString()));

                                    if (localCurrentTypesCache.HashTableKeyExist(objType))
                                    {
                                        objType = localCurrentTypesCache.HashTableGetKey(objType);
                                        instance.GetType().GetMethod("Add").Invoke(instance, new[] { objType });
                                        PrepareObjectToSerialize(objType);
                                    }
                                    else
                                    {
                                        instance.GetType().GetMethod("Add").Invoke(instance, new[] { objType });
                                        PrepareObjectToSerialize(objType);
                                    }
                                }
                            }
                            var arr = (Array)instance.GetType().GetMethod("ToArray").Invoke(instance, new object[] { });

                            prop.SetValue(obj, arr);
                        }
                        else
                        {
                            if (prop.PropertyType.IsArray && prop.PropertyType != (new object[] { }).GetType())
                            {
                                Type lst = typeof(List<>);
                                var constructedListType = lst.MakeGenericType(prop.PropertyType.GetElementType());

                                var instance = Activator.CreateInstance(constructedListType);
                                if (asmTypes.Count(t => t.Value.GetType() == instance.GetType()) > 0)
                                {
                                    //   asmTypes = asmTypes.OrderBy(t => Convert.ToInt32(t.Key.Substring(t.Key.LastIndexOf(']')+1))).Select(t => t).ToDictionary(t => t.Key, t => t.Value);
                                    var arrObj = asmTypes.First(t => t.Value.GetType() == instance.GetType());
                                    var val = arrObj.Value;
                                    // if (asmTypes.Count(t => t.Value.GetType() == instance.GetType()) > 1)
                                    asmTypes.Remove(arrObj.Key);
                                    var y = val.GetType().GetMethod("ToArray").Invoke(val, new object[] { });
                                    prop.SetValue(obj, y);
                                }
                                else
                                { 
                                
                                }
                            }
                            else
                            {
                                var objType = Activator.CreateInstance(prop.PropertyType);
                                //if (localCurrentTypesCache.HashTableKeyExist(objType))
                                //{
                                //    objType = localCurrentTypesCache.HashTableGetKey(objType);
                                //}
                                prop.SetValue(obj, objType);
                            }
                            if (prop.GetValue(obj) != null && prop.PropertyType != typeof(string))
                                PrepareObjectToSerialize(prop.GetValue(obj));
                        }
                    }
                    else
                    {
                        if (prop.PropertyType != typeof(string) && prop.PropertyType != typeof(int) && prop.PropertyType != typeof(bool) && prop.PropertyType != typeof(decimal)
                               && prop.PropertyType != typeof(double) && prop.PropertyType != typeof(float) && prop.PropertyType != typeof(long) && prop.PropertyType != typeof(DateTime)
                               && prop.PropertyType != typeof(char)
                               && !prop.PropertyType.IsArray)
                        {
                            PrepareObjectToSerialize(prop.GetValue(obj));
                        }
                    }
                }
            }
        }


        public void doOrdering(XmlNode node)
        {
            if (node.HasChildNodes)
            {
                //foreach (XmlNode ele in node.ChildNodes)
                //{
                //    doOrdering(ele);
                //}
                for (int i = 0; i < node.ChildNodes.Count; i++)
                {
                    doOrdering(node.ChildNodes.Item(i));
                }
            }

            XmlNode ans = null;
            XmlNode originalNode = null;
            GetElementByTagSuffix(node.LocalName, xmlSampleDoc.DocumentElement, ref ans);
            if (ans != null)
            {
                originalNode = ans;
            }
            if (originalNode != null)
            {
                if (node.PreviousSibling != null)
                {
                    if (node.PreviousSibling.LocalName != node.LocalName)
                    {
                        if (originalNode.PreviousSibling != null)
                        {
                            if (node.PreviousSibling.LocalName != originalNode.PreviousSibling.LocalName)
                            {
                                XmlNode parent = node.PreviousSibling.ParentNode;
                                XmlNode nodeTomove = null;
                                GetElementByTagSuffix(originalNode.PreviousSibling.LocalName, parent, ref nodeTomove);
                                parent.RemoveChild(nodeTomove);
                                parent.InsertBefore(nodeTomove, node);
                            }
                        }
                        else
                        {
                            Console.WriteLine(node.LocalName);
                            XmlNode parent = node.ParentNode;
                            parent.RemoveChild(node);
                            parent.PrependChild(node);
                        }
                    }
                }
            }
        }

        public void GetElementByTagSuffix(string suff, System.Xml.XmlNode node, ref System.Xml.XmlNode ans)
        {
            if (ans != null)
                return;
            System.Xml.XmlNodeList ds = node.ChildNodes;
            foreach (System.Xml.XmlNode c in ds)
            {
                if (c.LocalName.EndsWith(suff))
                {
                    ans = c;
                    break;
                }
                GetElementByTagSuffix(suff, c, ref ans);
            }
        }
    }
}
