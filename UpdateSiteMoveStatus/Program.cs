using System;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using System.Net;
using System.IO;

namespace UpdateSiteMoveStatus
{
    class Program
    {
        public static string FilePath { get; set; }
        public static string SrcPath { get; set; }
        public static string FileShortName { get; set; }
        public static string SiteUrl { get; set; }

        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("evict <JavaScript File> <Site Collection Url>");
            }


            if (System.IO.File.Exists(args[0]))
            {
                var file = new FileInfo(args[0]);
                FilePath = file.FullName;
                SrcPath = Path.Combine(file.Directory.FullName, @"..\..\", file.Name.Replace(".js", ".ts"));
                FileShortName = file.Name;
            }
            else
            {
                Console.WriteLine("You must specify a valid JavaScript file.");
                Console.ReadKey();
                Environment.Exit(1);
            }


            if (!string.IsNullOrEmpty(args[1]))
            {
                SiteUrl = args[1];
            }

            Startup();
            Console.WriteLine("Press the any key to continue!");
            Console.ReadKey();
        }

        private static void Startup()
        {
            using (var ctx = Context)
            {
                Site site = ctx.Site;

                // Get List
                var lists = ctx.LoadQuery(site.RootWeb.Lists.Where(l => l.RootFolder.Name == "SiteAssets"));
                try
                {
                    ctx.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message}\r\n{ex.StackTrace}", ConsoleColor.DarkRed);
                    Console.ReadKey();
                    Environment.Exit(1);
                }
                List siteAssets = lists.FirstOrDefault();

                if (siteAssets != null)
                {
                    // Check if file exists and check out if necessary
                    ctx.Load(site.RootWeb, w => w.ServerRelativeUrl);
                    try
                    {
                        ctx.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"{ex.Message}\r\n{ex.StackTrace}", ConsoleColor.DarkRed);
                        Console.ReadKey();
                        Environment.Exit(1);
                    }
                    var srcShortName = FileShortName.Replace(".js", ".ts");

                    var jsFile = site.RootWeb.GetFileByServerRelativeUrl($"{site.RootWeb.ServerRelativeUrl.TrimEnd('/')}/SiteAssets/{FileShortName}");
                    var srcMap = site.RootWeb.GetFileByServerRelativeUrl($"{site.RootWeb.ServerRelativeUrl.TrimEnd('/')}/SiteAssets/{FileShortName}.map");
                    var src = site.RootWeb.GetFileByServerRelativeUrl($"{site.RootWeb.ServerRelativeUrl.TrimEnd('/')}/SiteAssets/{srcShortName}");

                    ctx.Load(jsFile, f => f.Exists, f => f.CheckOutType);
                    ctx.Load(srcMap, f => f.Exists, f => f.CheckOutType);
                    ctx.Load(src, f => f.Exists, f => f.CheckOutType);

                    try
                    {
                        try
                        {
                            ctx.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"{ex.Message}\r\n{ex.StackTrace}", ConsoleColor.DarkRed);
                            Console.ReadKey();
                        }

                        if (jsFile.Exists && jsFile.CheckOutType == CheckOutType.None)
                        {
                            jsFile.CheckOut();
                            
                            Console.WriteLine($"{FileShortName} has been checked out!");
                        }

                        if (srcMap.Exists && srcMap.CheckOutType == CheckOutType.None)
                        {
                            srcMap.CheckOut();
                            
                            Console.WriteLine($"{FileShortName}.map has been checked out!");
                        }

                        if (src.Exists && src.CheckOutType == CheckOutType.None)
                        {
                            src.CheckOut();
                            
                            Console.WriteLine($"{FileShortName} has been checked out!");
                        }

                        try
                        {
                            ctx.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"{ex.Message}\r\n{ex.StackTrace}", ConsoleColor.DarkRed);
                            Console.ReadKey();
                            Environment.Exit(1);
                        }
                    }
                    catch (ServerException ex)
                    {
                        // expected if the file doesn't exist
                        if (ex.ServerErrorTypeName != "System.IO.FileNotFoundException")
                        {
                            throw;
                        }
                    }

                    // Get File Contents
                    string jsFileContents = System.IO.File.ReadAllText(FilePath);
                    string mapContents = System.IO.File.ReadAllText(FilePath + ".map");
                    string srcContents = System.IO.File.ReadAllText(SrcPath);

                    // Upload Script File
                    var jsFile2 = siteAssets.RootFolder.Files.Add(new FileCreationInformation
                    {
                        Content = Encoding.UTF8.GetBytes(jsFileContents),
                        Overwrite = true,
                        Url = FileShortName
                    });

                    var srcMap2 = siteAssets.RootFolder.Files.Add(new FileCreationInformation
                    {
                        Content = Encoding.UTF8.GetBytes(mapContents),
                        Overwrite = true,
                        Url = FileShortName + ".map"
                    });

                    var src2 = siteAssets.RootFolder.Files.Add(new FileCreationInformation
                    {
                        Content = Encoding.UTF8.GetBytes(srcContents),
                        Overwrite = true,
                        Url = FileShortName.Replace(".js", ".ts")
                    });

                    Console.WriteLine($"Files have been added!");

                    //if (jsFile.CheckOutType != CheckOutType.None)
                    //{
                    //    jsFile2.CheckIn(String.Empty, CheckinType.MajorCheckIn);
                    //    Console.WriteLine($"{FileShortName} has been checked in!");
                    //}

                    //if (srcMap.CheckOutType != CheckOutType.None)
                    //{
                    //    srcMap2.CheckIn(String.Empty, CheckinType.MajorCheckIn);
                    //    Console.WriteLine($"{FileShortName}.map has been checked in!");
                    //}

                    //if (src.CheckOutType != CheckOutType.None)
                    //{
                    //    src2.CheckIn(String.Empty, CheckinType.MajorCheckIn);
                    //    Console.WriteLine($"{srcShortName} has been checked in!");
                    //}

                    try
                    {
                        ctx.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"{ex.Message}\r\n{ex.StackTrace}", ConsoleColor.DarkRed);
                        Console.ReadKey();
                        Environment.Exit(1);
                    }
                    Console.WriteLine($"Server synced!");


                    // Register Custom Action
                    UserCustomAction customAction = site.UserCustomActions.Add();
                    customAction.Location = "ScriptLink";
                    customAction.ScriptSrc = $"~SiteCollection/SiteAssets/{FileShortName}";
                    customAction.Sequence = 1000;
                    customAction.Update();
                    try
                    {
                        ctx.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"{ex.Message}\r\n{ex.StackTrace}", ConsoleColor.DarkRed);
                        Console.ReadKey();
                        Environment.Exit(1);
                    }
                }
                else
                {
                    throw new FileNotFoundException("'SiteAssets' not found");
                }
            }

        }
        private static ClientContext _context;

        public static ClientContext Context
        {
            get {
                if (_context != null)
                {
                    return _context;
                }
                else
                {
                    var ctx = new ClientContext(SiteUrl)
                    {
                        Credentials = CredentialCache.DefaultNetworkCredentials
                    };
                    return _context = ctx;
                }
            }
        }
    }
}
