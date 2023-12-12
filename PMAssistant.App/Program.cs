using Microsoft.Extensions.Configuration;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System;
using System.Linq;
using System.IO;
using System.Threading;

namespace PMAssistant.App
{
    class Program
    {
        static IConfigurationRoot configuration;
        static string CommonTemplatesPath;
        static string ProjectDirectory;
        static Product Product;
        static List<Product> Products;
        static List<Milestone> Milestones;
        static Dictionary<string, string> Variables;
        static string[] BasicDirectories;
        static int MoMFilesCount;

        #region Constants
        const int DEFAULT_MoM_FILES_COUNT = 8;
        const int MAX_MoM_FILES_COUNT = 12;
        //const int MAX_MILESTONES_COUNT = 8;
        const string WCF_SHEET_NAME = "Form";
        const string PreSalesDirectoryName = "Pre-Sales";
        const string MinutesDirectoryName = "Minutes";
        const string WCFDirectoryName = "WCF";
        #endregion
        static void Main(string[] args)
        {
            Initialize();

            //var PMName = LoadPMDetails();//considered to collect PM Data from Active Directory..failed perhaps for some policy by the organization
            Console.WriteLine($"Hi! May the client be cooperative and your forcasts be accurate!");
            Console.WriteLine($"Let's get started, This gonna be quick!{Environment.NewLine}");

            Console.WriteLine($"What is the Product? Please type the name of the product from the following list:{Environment.NewLine}");
            Console.WriteLine(string.Join(Environment.NewLine, (from x in Products select x.Name).ToList()));
            while (true)
            {
                var input = Console.ReadLine();
                Product = Products.Where(x => x.Name.ToLower() == input.ToLower()).SingleOrDefault();
                if (Product == null)
                {
                    Console.WriteLine("Please provide a product from the list above..");
                }
                else
                {
                    Variables.Add("#Product#", Product.Name);
                    break;
                }
            }
            Console.WriteLine("Please specify the Project directory path:");
            while (true)
            {
                ProjectDirectory = Console.ReadLine();
                if (!string.IsNullOrEmpty(ProjectDirectory))
                {
                    var validDirectory = new DirectoryInfo(ProjectDirectory).Exists;
                    if (validDirectory)
                        break;
                }
                Console.WriteLine("Please provide a valid path..");
            }
            if (!Directory.Exists(Path.Combine(ProjectDirectory, PreSalesDirectoryName)))
            {
                Console.WriteLine("Please specify the Project [Pre-Sales] directory path:");
                while (true)
                {
                    var path = Console.ReadLine();
                    if (string.IsNullOrEmpty(path))
                    {
                        Console.WriteLine("Will create an empty Pre-Sales Directory for you, kindly consider to fill in related data later..");
                        CreateDirectory(PreSalesDirectoryName);
                        break;
                    }
                    if (!string.IsNullOrEmpty(path))
                    {
                        var validDirectory = new DirectoryInfo(path).Exists;
                        if (validDirectory)
                        {
                            CreatePreSalesDirctory(path);
                            break;
                        }
                    }
                    Console.WriteLine("Please provide a valid path..");
                }
            }
            var variablesFilePath = CreateVariablesFile();
            Console.WriteLine("------------------------");
            Console.WriteLine("Please fill the variables in the Excel workbook we've just opened for you");
            Console.WriteLine($"File didnt open? You can open it manually at this path: {Environment.NewLine}{variablesFilePath}{Environment.NewLine}");
            Console.WriteLine($"Please make sure to save your changes, close the file and hit enter...");
            Console.ReadLine();
            while (true)
            {
                if (IsFileinUse(variablesFilePath))
                {
                    Console.WriteLine($"File is still open, please close the file and hit enter..");
                    Console.ReadLine();
                    continue;
                }
                var noMissingInfo = LoadVariablesFile(variablesFilePath);
                if (noMissingInfo)
                {
                    break;
                }
                Console.WriteLine($"You missed some variables {Environment.NewLine} We opened the file again, please fill missing variables");
                Console.WriteLine($"Once complete please hit enter to proceed..");
                Console.ReadLine();
            }

            Console.WriteLine($"How many Minutes of Meeting files would you like to create?");
            Console.WriteLine($"Please provide a number (Maximum of {MAX_MoM_FILES_COUNT}) or hit enter to keep the default of {DEFAULT_MoM_FILES_COUNT}");
            while (true)
            {
                var input = Console.ReadLine();
                if (string.IsNullOrEmpty(input))
                {
                    MoMFilesCount = DEFAULT_MoM_FILES_COUNT;
                    break;
                }
                else
                {
                    if (int.TryParse(input, out int desiredFilesCount))
                    {
                        if (desiredFilesCount > MAX_MoM_FILES_COUNT)
                        {
                            Console.WriteLine($"Sorry you exceeded the max nb of MoM files, we will revert to the Max value instead..");
                            MoMFilesCount = MAX_MoM_FILES_COUNT;
                        }
                        else
                        {
                            MoMFilesCount = desiredFilesCount;
                        }
                        break;
                    }
                    Console.WriteLine("Please provide a valid number..");
                }
            }
            Console.WriteLine("That's it, Thank you! Process will start now, will inform you once complete!");
            Console.WriteLine("------------------------");

            foreach (var directoryName in BasicDirectories)
            {
                Console.WriteLine($"Setting up [{directoryName}] Directory..");
                CreateDirectoryAndTemplates(directoryName);
                Console.WriteLine("------------------------");
            }

            Console.WriteLine("Generating Minutes Directory..");
            GenerateMinutesFiles();
            Console.WriteLine("------------------------");

            Console.WriteLine("Setting up WCF Directory..");
            GenerateWCFFiles();

            Console.WriteLine("------------------------");
            Console.WriteLine("Press any key to exit..");
            Console.ReadLine();
        }

        #region Init
        private static void Initialize()
        {
            var userDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            Variables = new Dictionary<string, string> { };
            configuration = new ConfigurationBuilder().SetBasePath(Directory.GetParent(AppContext.BaseDirectory).FullName).AddJsonFile("appsettings.json", false).Build();
            var productsTemplatesPath = configuration.GetSection("ProductsTemplatesPath").Value;
            CommonTemplatesPath = configuration.GetSection("CommonTemplatesPath").Value;
            BasicDirectories = configuration.GetSection("BasicDirectories").GetChildren().ToArray().Select(c => c.Value).ToArray();

            productsTemplatesPath = $"{userDirectory}\\{productsTemplatesPath}";
            CommonTemplatesPath = $"{userDirectory}\\{CommonTemplatesPath}";
            LoadProducts(productsTemplatesPath);
        }
        private static void LoadProducts(string productsTemplatesPath)
        {
            Products = new List<Product> { };

            var directories = Directory.GetDirectories(productsTemplatesPath);
            foreach (var directory in directories)
            {
                var name = new DirectoryInfo(directory).Name;
                if (!name.StartsWith("_"))
                    Products.Add(new Product { Name = name, TemplatesDirectory = directory });
            }
        }
        //private static string LoadPMDetails()
        //{
        //    var name = string.Empty;
        //    var emailAddress = string.Empty;
        //    try
        //    {
        //        PrincipalContext ctx = new PrincipalContext(ContextType.Domain);
        //        UserPrincipal user = UserPrincipal.Current;
        //        name = user.DisplayName;
        //        emailAddress = user.EmailAddress;
        //    }
        //    catch (Exception ex)
        //    {
        //        name = string.Empty;
        //        emailAddress = string.Empty;
        //    }
        //    if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(emailAddress))
        //    {
        //        Console.WriteLine("Unable to fech your account information..");
        //        Console.WriteLine("Please enter your name:");
        //        name = Console.ReadLine();
        //        Console.WriteLine("and your email address:");
        //        emailAddress = Console.ReadLine();
        //    }
        //    Variables.Add("#PM Name#", name);
        //    Variables.Add("#PM Email#", emailAddress);
        //    return name;
        //}
        #endregion

        #region Core
        static bool IsGenericDirectory(string identifier)
        {
            return identifier == PreSalesDirectoryName || identifier == WCFDirectoryName || identifier == MinutesDirectoryName; ;
        }
        static void CreateDirectoryAndTemplates(string identifier)
        {
            //var path = CreateDirectory(identifier); //create if not exists
            var path = Path.Combine(ProjectDirectory, identifier);
            if (Directory.Exists(path))
            {
                Console.WriteLine($"Director {identifier} exists already --> No files generated!!");
                return;
            }
            Directory.CreateDirectory(path);

            if (IsGenericDirectory(identifier))
                return;

            var sourceFilePath = GetTemplatePath(Product.TemplatesDirectory, identifier);
            if (string.IsNullOrEmpty(sourceFilePath))
            {
                Console.WriteLine($"No {Product.Name}-specific {identifier} template available");
                Console.WriteLine($"Checking for generic {identifier} template..");
                sourceFilePath = GetTemplatePath(CommonTemplatesPath, identifier);
                if (string.IsNullOrEmpty(sourceFilePath))
                {
                    Console.WriteLine($"No generic {identifier} template available as well..");
                    return;
                }
            }
            var fileName = new DirectoryInfo(sourceFilePath).Name;
            var destinationFilePath = Path.Combine(path, SetVariablesValue(fileName));
            if (!File.Exists(destinationFilePath))
                File.Copy(sourceFilePath, destinationFilePath);

        }
        private static string GetTemplatePath(string directoryPath, string identifier, string type = "")
        {
            var files = Directory.GetFiles(directoryPath);
            foreach (var file in files)
            {
                if (file.ToLower().Contains(identifier.ToLower()) && (string.IsNullOrEmpty(type) || (!string.IsNullOrEmpty(type) && file.ToLower().Contains(type.ToLower()))))
                    return file;
            }
            return string.Empty;
        }
        #endregion

        #region PreSales
        static void CreatePreSalesDirctory(string sourceDir)
        {
            var destinationDir = Path.Combine(ProjectDirectory, PreSalesDirectoryName);
            //CreateShortcut(sourceDir, PreSalesDirectory); //considered to create a shorcut to the Presales directory instead of Copy to avoid duplication,
            //                                                failed..didnt have the chance to investigate
            CopyDirectory(sourceDir, destinationDir);
        }
        #endregion

        #region Minutes
        static void GenerateMinutesFiles()
        {
            //var path = CreateDirectory(MinutesDirectoryName);
            var path = Path.Combine(ProjectDirectory, MinutesDirectoryName);
            if (!Directory.Exists(path))
            {
                Console.WriteLine("Minutes Not enabled!");
                return;
            }
            var sourceFilePath = GetTemplatePath(CommonTemplatesPath, MinutesDirectoryName);
            var fileName = Path.GetFileName(sourceFilePath);
            DateTime nextWeeklyCall = GetNextWeeklyCall();

            Console.WriteLine($"Creating {MoMFilesCount} Minutes of Meeting files starting {nextWeeklyCall.ToString("dd-MM-yyyy")}..");

            for (int i = 0; i < MoMFilesCount; i++)
            {
                var date = nextWeeklyCall.AddDays(7 * i).ToString("dd-MMM-yyyy");
                Variables.Add("#dd-MMM-yyyy#", date);
                CreateWordDocument(sourceFilePath, path);
                Variables.Remove("#dd-MMM-yyyy#");
                Console.WriteLine($"--> [{date}] Minutes of Meeting file created!");
            }
        }
        private static DateTime GetNextWeeklyCall()
        {
            var dateTime = DateTime.Now;
            if (!DateTime.TryParse(SetVariablesValue("#Next Weekly Call#"), out dateTime))
            {
                Console.WriteLine("WARNING, Next weekly call variable not provided, or has invalid value, setting the next weekly call in one week..");
                dateTime = DateTime.Now;
            }
            return dateTime;
        }
        #endregion

        #region WCF
        static void GenerateWCFFiles()
        {
            var path = Path.Combine(ProjectDirectory, WCFDirectoryName);
            if (!Directory.Exists(path))
            {
                Console.WriteLine("WCF Not enabled!");
                return;
            }
            //var path = CreateDirectory(WCFDirectoryName);
            var sourceFilePath = GetTemplatePath(CommonTemplatesPath, "Work Completion Form");
            var fileName = Path.GetFileName(sourceFilePath);

            LoadMilestones();
            foreach (var milestone in Milestones)
            {
                Variables.Add("#NO#", milestone.Id.ToString());
                Variables.Add("#Milestone Description#", milestone.Description);
                Variables.Add("#Milestone Full Description#", milestone.FullDescription);

                var destinationFilePath = Path.Combine(path, SetVariablesValue(fileName));
                File.Copy(sourceFilePath, destinationFilePath);

                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(destinationFilePath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
                try
                {
                    Excel.Sheets excelSheets = excelWorkbook.Worksheets;
                    Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(WCF_SHEET_NAME);
                    Excel.Range excelCell = excelWorksheet.get_Range("A1", "Z100");
                    foreach (Excel.Range cell in excelCell.Cells)
                    {
                        if (ValueContainsVariable(cell.Value2))
                        {
                            cell.Value2 = SetVariablesValue(cell.Value2);
                        }
                    }
                    excelWorkbook.Save();
                    Console.WriteLine($"--> [Milestone {milestone.Id}] Excel file created!");

                    Variables.Remove("#NO#");
                    Variables.Remove("#Milestone Description#");
                    Variables.Remove("#Milestone Full Description#");
                }
                catch
                (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                excelWorkbook.Close();
            }
            GenerateWCFEmails(path);
            ExportExcelToPDF(path);
        }
        static void GenerateWCFEmails(string directory)
        {
            var emailTemplatePath = GetTemplatePath(CommonTemplatesPath, "Work Completion Email");
            var fullName = GetVariableValue("#Client PM Name#");
            Variables.Add("#Client PM First Name#", fullName.Split(' ')[0]);
            foreach (var milestone in Milestones)
            {
                Variables.Add("#NO#", milestone.Id.ToString());
                Variables.Add("#Milestone Description#", milestone.Description);
                CreateWordDocument(emailTemplatePath, directory);
                Console.WriteLine($"--> [Milestone {milestone.Id}] Email body created!");
                Variables.Remove("#NO#");
                Variables.Remove("#Milestone Description#");
            }
        }
        private static void LoadMilestones()
        {
            Milestones = new List<Milestone> { };
            var milestones = Variables.Where(x => x.Key.StartsWith("#Milestone")
                                                && !x.Value.StartsWith("<")
                                                && !string.IsNullOrEmpty(x.Value)).Select(x => x.Value).ToList();
            int i = 1;
            foreach (var milestone in milestones)
            {
                var fullDescription = string.Empty;
                var description = milestone;
                if (description.StartsWith("Milestone"))
                {
                    fullDescription = description;
                    description = CleanMilestoneName(description);
                }
                else
                {
                    fullDescription = $"Milestone {i}: {description}";
                }
                Milestones.Add(new Milestone { Id = i++, Description = description, FullDescription = fullDescription });
            }
        }
        private static string CleanMilestoneName(string description)
        {
            if (description.Contains(':'))
                description = description.Split(':')[1];
            else
            {
                description = description.Split("Milestone")[1].Trim();
                if (description.Length > 2)
                {
                    string leadingCharacters = description.Substring(0, 2);
                    var noiseCharacters = new List<string> { "1 ", "2 ", "3 ", "I ", "II", "III", "IV" };
                    if (noiseCharacters.Any(x => x == leadingCharacters))
                        description = description.Substring(2);
                }
            }
            return description;
        }
        #endregion

        #region Variables
        static string CreateVariablesFile()
        {
            var sourceFilePath = "Resources/Variables.csv";
            var destinationFilePath = Path.Combine(ProjectDirectory, $"Variables.csv");
            File.Copy(sourceFilePath, destinationFilePath);

            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.OpenText(destinationFilePath, Comma: true);
            return destinationFilePath;
        }
        static bool LoadVariablesFile(string path)
        {
            var valid = true;

            var lines = File.ReadAllLines(path);
            foreach (var line in lines)
            {
                var values = line.Split(',');
                if (VariableMissing(values[0], values[1]))
                {
                    valid = false;
                    break;
                }
                Variables.Add(values[0], values[1]);
            }
            if (valid)
            {
                //All variables have been provided by user and now loaded into Variables Dictionnary, we can delete the file
                File.Delete(path);
            }
            else
            {
                //Reopen the file for user to provide missing info + Reset Variables dictionnary
                var excelApp = new Excel.Application();
                excelApp.Visible = true;
                excelApp.Workbooks.OpenText(path, Comma: true);
                Variables = new Dictionary<string, string> { };
            }
            return valid;
        }
        private static bool VariableMissing(string key, string value)
        {
            if (string.IsNullOrEmpty(key) || key.StartsWith("#M"))
                return false;

            return (string.IsNullOrEmpty(value) || value.ToLower().Contains("please fill"));
        }
        private static bool ValueContainsVariable(dynamic value)
        {
            return (value != null && value.GetType() == typeof(string) && ((string)value).Contains('#'));
        }
        private static string GetVariableValue(dynamic value)
        {
            if (Variables.ContainsKey((string)value))
            {
                var variable = Variables[(string)value];
                if (string.IsNullOrEmpty(variable))
                    return "<Not Available>";
                return variable;
            }
            return string.Empty;
        }
        private static string SetVariablesValue(string value)
        {
            foreach (var item in Variables)
            {
                value = value.Replace(item.Key, item.Value);
            }
            return value;
        }
        #endregion

        #region Documents
        static bool IsFileinUse(string filePath)
        {
            var file = new FileInfo(filePath);
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }
        static string CreateDirectory(string category)
        {
            var path = Path.Combine(ProjectDirectory, category);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }
        static void CreateWordDocument(string templatePath, string destinationDirectory)
        {
            var fileName = Path.GetFileName(templatePath);
            Word._Application Application = new Word.Application();
            Application.Visible = false;
            Application.ScreenUpdating = false;
            Application.WindowState = Word.WdWindowState.wdWindowStateMinimize;
            var destinationEmailPath = Path.Combine(destinationDirectory, SetVariablesValue(fileName));
            File.Copy(templatePath, destinationEmailPath);
            Application.Documents.Open(destinationEmailPath);
            SearchReplace(Application);
            Application.Documents.Close();
            Thread.Sleep(1000);
        }
        private static void SearchReplace(Word._Application Application)
        {
            object missing = System.Reflection.Missing.Value;

            Word.Find findObject = Application.Selection.Find;
            foreach (var item in Variables)
            {
                findObject.ClearFormatting();
                findObject.Text = item.Key;
                findObject.Replacement.ClearFormatting();
                findObject.Replacement.Text = item.Value;

                object replaceAll = Word.WdReplace.wdReplaceAll;
                findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref replaceAll, ref missing, ref missing, ref missing, ref missing);
            }
        }
        static void ExportExcelToPDF(string directory)
        {
            int i = 0;
            var files = Directory.GetFiles(directory, "*.xlsx");
            foreach (var file in files)
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(file,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
                object outputFileName = excelWorkbook.FullName.Replace(".xlsx", ".pdf");
                excelWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputFileName);
                Console.WriteLine($"--> [Milestone {++i}] PDF file created!");
                excelWorkbook.Close();
                Thread.Sleep(1000);
            }
        }
        static void CopyDirectory(string sourceDir, string destinationDir, bool recursive = true)
        {
            // Get information about the source directory
            var dir = new DirectoryInfo(sourceDir);

            // Check if the source directory exists
            if (!dir.Exists)
                throw new DirectoryNotFoundException($"Source directory not found: {dir.FullName}");

            // Cache directories before we start copying
            DirectoryInfo[] dirs = dir.GetDirectories();

            // Create the destination directory
            Directory.CreateDirectory(destinationDir);

            // Get the files in the source directory and copy to the destination directory
            foreach (FileInfo file in dir.GetFiles())
            {
                string targetFilePath = Path.Combine(destinationDir, file.Name);
                file.CopyTo(targetFilePath);
            }

            // If recursive and copying subdirectories, recursively call this method
            if (recursive)
            {
                foreach (DirectoryInfo subDir in dirs)
                {
                    string newDestinationDir = Path.Combine(destinationDir, subDir.Name);
                    CopyDirectory(subDir.FullName, newDestinationDir, true);
                }
            }
        }
        //static void CreateShortcut(string sourcePath, string shortcutName, string destinationPath="")
        //{
        //    if (string.IsNullOrEmpty(destinationPath))
        //        destinationPath = ProjectDirectory;
        //    IWshRuntimeLibrary.IWshShortcut shortcut;
        //    IWshRuntimeLibrary.IWshShell_Class wshShell = new IWshRuntimeLibrary.IWshShell_Class();
        //    shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(sourcePath);
        //    shortcut.TargetPath = Path.Combine(destinationPath, shortcutName);
        //    shortcut.Save();
        //}
        #endregion
    }
    struct Milestone
    {
        public int Id { get; set; }
        public string Description { get; set; }
        public string FullDescription { get; set; }
    }
    class Product
    {
        public string Name { get; set; }
        public string TemplatesDirectory { get; set; }
    }

}
