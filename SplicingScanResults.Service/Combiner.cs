using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using MigraDoc.DocumentObjectModel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using Topshelf;
using ZXing;

namespace SplicingScanResults.Service
{
    public class Combiner : ServiceControl, IDisposable
    {
        private Regex validator;
        private FileSystemWatcher watcher;
        private BarcodeReader barcodeReader = new BarcodeReader();
        private ManualResetEventSlim workToDo = new ManualResetEventSlim(false);
        private ManualResetEventSlim fileSynchronizer = new ManualResetEventSlim(true);
        private CancellationTokenSource tokenSource;
        //private ConcurrentQueue<FileSystemEventArgs> fileQueue = new ConcurrentQueue<FileSystemEventArgs>();

        private FileSystemEventArgs eventArg; 

        public TimeSpan Timeout { get; set; }
        public string RootDirectory { get; }


        public Combiner(string scanerRootPath)
        {
            if (!Directory.Exists(scanerRootPath))
            {
                Directory.CreateDirectory(scanerRootPath);
            }

            RootDirectory = scanerRootPath;

            validator = new Regex(@"^img_(\d{3})\.(jpg|png|bmp)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

            watcher = new FileSystemWatcher(RootDirectory)
            {
                IncludeSubdirectories = false,
                EnableRaisingEvents = false
            };

            watcher.Created += (source, e) => 
                {
                    workToDo.Set();
                    eventArg = e;
                };

            Timeout = TimeSpan.FromSeconds(500);
        }

        public bool Start(HostControl hostControl)
        {
            watcher.EnableRaisingEvents = true;
            tokenSource = new CancellationTokenSource();

            Task.Run(() =>
            {
                Match match;
                Result result;
                HashSet<string> listOfImages = new HashSet<string>();
                int prevNumber = -1;

                while (!tokenSource.IsCancellationRequested)
                {
                    if (workToDo.Wait(Timeout))
                    {
                        workToDo.Reset();

                        match = validator.Match(eventArg.Name);

                        if (match.Success)
                        {
                            var imageNumber = int.Parse(match.Groups[1].Value);

                            if (prevNumber == -1 || imageNumber - prevNumber == 1)
                            {
                                try
                                {
                                    Task.Delay(TimeSpan.FromMilliseconds(100)).Wait();
                                    fileSynchronizer.Wait();
                                    using (var imageStream = new FileStream(eventArg.FullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                    using (var image = new Bitmap(imageStream))
                                    {
                                        result = barcodeReader.Decode(image);
                                    }
                                    fileSynchronizer.Set();

                                    prevNumber = imageNumber;

                                    listOfImages.Add(eventArg.FullPath);

                                    if (result != null)
                                    {
                                        InitTask(listOfImages);

                                        prevNumber = -1;
                                    }
                                }
                                catch (IOException ex) when (ex.Message.Contains("because it is being used by another process"))
                                {
                                    Console.WriteLine(ex.Message);
                                }
                            }
                            else
                            {
                                InitTask(listOfImages);

                                prevNumber = imageNumber;

                                listOfImages.Add(eventArg.FullPath);
                            }
                        }
                        else
                        {
                            File.Delete(eventArg.FullPath);
                        }

                    }
                    else
                    {
                        if (listOfImages.Count > 0)
                        {
                            InitTask(listOfImages);
                        }

                        prevNumber = -1;
                    }
                }
            }, tokenSource.Token);

            return true;
        }

        public bool Stop(HostControl hostControl)
        {
            watcher.EnableRaisingEvents = false;

            tokenSource.Cancel();

            return true;
        }

        private void InitTask(HashSet<string> imageList)
        {
            Task pdfCreationTask = new Task(images => CreatePDF((List<string>)images),
                new List<string>(imageList));

            pdfCreationTask.Start();
            imageList.Clear();
        }

        private void CreatePDF(List<string> images)
        {
            var resultDirectory = Directory.CreateDirectory(Path.Combine(RootDirectory, "Result", DateTime.Today.ToString("dd-MM-yy")));
            var pdfName = Path.Combine(resultDirectory.FullName, $"{Guid.NewGuid().ToString()}.pdf");

            int i = 0;
            try
            {
                PdfDocument document = new PdfDocument();

                for (; i < images.Count; i++)
                {
                    var page = document.AddPage();
                    using (var stream = new FileStream(images[i], FileMode.Open))
                    {
                        XGraphics gfx = XGraphics.FromPdfPage(page);

                        gfx.DrawImage(XImage.FromStream(stream), new XRect(0, 0, page.Width, page.Height));
                    }
                }

                document.Save(pdfName);
                document.Dispose();

                fileSynchronizer.Wait();
                images.ForEach(File.Delete);
                fileSynchronizer.Set();
            }
            catch(Exception ex)
            {
                fileSynchronizer.Set();

                Console.WriteLine(ex.Message);

                images.Add(pdfName);

                OnError(images, i, ex.Message);

                resultDirectory.Delete(true);
            }
        }

        private void OnError(List<string> files, int index, string message)
        {
            var errorDirectory =  Directory.CreateDirectory(Path.Combine(RootDirectory, "Broken", DateTime.Today.ToString("dd-MM-yy"), Guid.NewGuid().ToString()));

            for (var i = 0; i < files.Count; i++)
            {
                var fileName = Path.GetFileName(files[i]);

                fileSynchronizer.Wait();
                File.Move(Path.Combine(RootDirectory, fileName), Path.Combine(errorDirectory.FullName, i == index ? $"{message}_{fileName}" : fileName));
                fileSynchronizer.Set();
            }
        }

        public void Dispose()
        {
            watcher?.Dispose();
        }
    }
}
