using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.IO;

namespace CreateIcons
{
    public partial class Form1 : Form
    {
        private static Bitmap bmpScreenshot;
        private static Graphics gfxScreenshot;

        // where to save icons
        public string iconPath = @"C:\icons\";
        
        // source directory
        public string sourceDir = Directory.GetCurrentDirectory() + "\\test";
        
        // final directory after cloning
        public string newDirectory = @"C:\new";
		
        public List<string> extensions = new List<string>();

        public List<string> files = new List<string>();
        public List<string> directories = new List<string>();

        public List<double> positions;
        public int currentFileIndex = 0;
        public int currentTimerPosition = 0;
        public string currentDirectoryName = "";

        public System.Timers.Timer timer = new System.Timers.Timer();
        public System.Timers.Timer playTimer = new System.Timers.Timer();
        public Form1()
        {
            InitializeComponent();

            extensions.Add("*.mp4");
            extensions.Add("*.mov");
            extensions.Add("*.avi");

            // media player cant handle either .mkv extension or really big files, haven't really tested it
            // extensions.Add("*.mkv");

            // min. 1000 ms is recommended, because under that media player might not be able to load in time
            timer.Interval = 1500;
            timer.Elapsed += timer_Elapsed;

            // also min. 1000 ms, this is for when a new movie starts in media player
            playTimer.Interval = 1000;
            playTimer.Elapsed += play_timer_Elapsed;

            axWindowsMediaPlayer1.settings.mute = true;
        }

        /// <summary>
        /// Get's all the movie files from sourceDir
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            richTextBox1.Hide();
            var folders = System.IO.Directory.GetDirectories(sourceDir, "*", System.IO.SearchOption.AllDirectories);

            if (folders.Length > 0)
            {
                foreach (string f in folders)
                {
                    directories.Add(System.IO.Path.GetFileName(f));

                    for (var i = 0; i < extensions.Count; i++)
                    {
                        string[] filesNotSavedYet = Directory.GetFiles(f, extensions[i]);

                        if (filesNotSavedYet.Length > 0)
                        {
                            foreach (var file in filesNotSavedYet)
                            {
                                files.Add(file);
                            }
                        }
                    }
                }
            }

            if (files.Count > 0)
            {
                startPlaying(files[currentFileIndex]);
            }
        }

        /// <summary>
        /// Starts playing the nth movie
        /// <param name="file">Path of the file to be played</param>
        /// </summary>
        public void startPlaying(string file)
        {
            currentDirectoryName = directories[currentFileIndex];
            Directory.CreateDirectory(currentDirectoryName);
            axWindowsMediaPlayer1.URL = file;
            axWindowsMediaPlayer1.Ctlcontrols.currentPosition = 0;
            playTimer.Start();
        }

        /// <summary>
        /// The media player in winform takes time to start a movie
        /// </summary>
        public void play_timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            positions = getDuration(13);
            axWindowsMediaPlayer1.Ctlcontrols.currentPosition = positions[currentTimerPosition];

            timer.Start();
            playTimer.Stop();
        }

        /// <summary>
        /// Divides the movie x times for getting timestamps
        /// Offset is distracted from every division(it's removeable, but then you have to remove times + 1)
        /// only purpose of this was, that I wanted icons from the end too
        /// </summary>
        /// <param name="times">Root of the directory</param>
        /// <returns>List of the durations</returns>
        public List<double> getDuration(int times)
        {
            List<double> list = new List<double>();
            double adder = axWindowsMediaPlayer1.Ctlcontrols.currentItem.duration / times;
            int offset = 15;

            for (int i = 1; i < times + 1; i++)
            {
                list.Add((adder * i) - offset);
            }

            return list;
        }

        /// <summary>
        /// Handles taking screenshots, starting new movies
        /// </summary>
        public void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            currentTimerPosition += 1;

            string[] fileNames = files[currentFileIndex].Split('\\');

            if (currentTimerPosition > 12)
            {
                // takeScreenshot(fileNames[fileNames.Length - 1] + "_" + (currentTimerPosition - 1));
                positions.Clear();
                timer.Stop();
                currentFileIndex++;
                currentTimerPosition = 0;
                if (files.Count - 1 < currentFileIndex)
                {
                    var directories = Directory.GetDirectories(Directory.GetCurrentDirectory());
                    foreach (var dir in directories)
                    {
                        if (dir == sourceDir)
                        {
                            continue;
                        }
                        Directory.CreateDirectory(iconPath + System.IO.Path.GetFileName(dir));
                        CloneDirectory(dir, iconPath + System.IO.Path.GetFileName(dir));
                        Directory.Delete(dir, true);
                    }
                    axWindowsMediaPlayer1.Dispose();
                    return;
                }
                startPlaying(files[currentFileIndex]);
                return;
            }
            else
            {
                takeScreenshot(fileNames[fileNames.Length - 1] + "_" + (currentTimerPosition - 1));
                axWindowsMediaPlayer1.Ctlcontrols.currentPosition = positions[currentTimerPosition];
            }
        }

        /// <summary>
        /// Clones a directory (subdirectories too) from root to destination
        /// Found this method online, can't remember who wrote it
        /// </summary>
        /// <param name="root">Root of the directory</param>
        /// <param name="dest">Destination directory</param>
        private static void CloneDirectory(string root, string dest)
        {
            foreach (var directory in Directory.GetDirectories(root))
            {
                string dirName = Path.GetFileName(directory);
                if (!Directory.Exists(Path.Combine(dest, dirName)))
                {
                    Directory.CreateDirectory(Path.Combine(dest, dirName));
                }
                CloneDirectory(directory, Path.Combine(dest, dirName));
            }

            foreach (var file in Directory.GetFiles(root))
            {
                File.Copy(file, Path.Combine(dest, Path.GetFileName(file)));
            }
        }

        /// <summary>
        /// Converts a PNG image to an icon (ico)
        /// </summary>
        /// <param name="input">The input stream</param>
        /// <param name="output">The output stream</param>
        /// <param name="size">Needs to be a factor of 2 (16x16 px by default)</param>
        /// <param name="preserveAspectRatio">Preserve the aspect ratio</param>
        /// <returns>Wether or not the icon was succesfully generated</returns>
        public static bool ConvertToIcon(Stream input, Stream output, int size = 16, bool preserveAspectRatio = false)
        {
            var inputBitmap = (Bitmap)Bitmap.FromStream(input);
            if (inputBitmap == null)
                return false;

            float width = size, height = size;
            if (preserveAspectRatio)
            {
                if (inputBitmap.Width > inputBitmap.Height)
                    height = ((float)inputBitmap.Height / inputBitmap.Width) * size;
                else
                    width = ((float)inputBitmap.Width / inputBitmap.Height) * size;
            }

            var newBitmap = new Bitmap(inputBitmap, new Size((int)width, (int)height));
            if (newBitmap == null)
                return false;

            // save the resized png into a memory stream for future use
            using (MemoryStream memoryStream = new MemoryStream())
            {
                newBitmap.Save(memoryStream, ImageFormat.Png);

                var iconWriter = new BinaryWriter(output);
                if (output == null || iconWriter == null)
                    return false;

                // 0-1 reserved, 0
                iconWriter.Write((byte)0);
                iconWriter.Write((byte)0);

                // 2-3 image type, 1 = icon, 2 = cursor
                iconWriter.Write((short)1);

                // 4-5 number of images
                iconWriter.Write((short)1);

                // image entry 1
                // 0 image width
                iconWriter.Write((byte)width);
                // 1 image height
                iconWriter.Write((byte)height);

                // 2 number of colors
                iconWriter.Write((byte)0);

                // 3 reserved
                iconWriter.Write((byte)0);

                // 4-5 color planes
                iconWriter.Write((short)0);

                // 6-7 bits per pixel
                iconWriter.Write((short)32);

                // 8-11 size of image data
                iconWriter.Write((int)memoryStream.Length);

                // 12-15 offset of image data
                iconWriter.Write((int)(6 + 16));

                // write image data
                // png data must contain the whole png data file
                iconWriter.Write(memoryStream.ToArray());

                iconWriter.Flush();
            }

            return true;
        }

        /// <summary>
        /// Converts a PNG image to an icon (ico)
        /// </summary>
        /// <param name="inputPath">The input path</param>
        /// <param name="outputPath">The output path</param>
        /// <param name="size">Needs to be a factor of 2 (16x16 px by default)</param>
        /// <param name="preserveAspectRatio">Preserve the aspect ratio</param>
        /// <returns>Wether or not the icon was succesfully generated</returns>
        public static bool ConvertToIcon(string inputPath, string outputPath, int size = 16, bool preserveAspectRatio = false)
        {
            using (FileStream inputStream = new FileStream(inputPath, FileMode.Open))
            using (FileStream outputStream = new FileStream(outputPath, FileMode.OpenOrCreate))
            {
                var x = ConvertToIcon(inputStream, outputStream, size, preserveAspectRatio);
                inputStream.Dispose();
                outputStream.Dispose();
                return x;
            }
        }

        /// <summary>
        /// Converts a PNG image to an icon (ico)
        /// </summary>
        /// <param name="inputPath">Image object</param>
        /// <param name="preserveAspectRatio">Preserve the aspect ratio</param>
        /// <returns>ico byte array / null for error</returns>
        public static byte[] ConvertToIcon(Image image, bool preserveAspectRatio = false)
        {
            MemoryStream inputStream = new MemoryStream();
            image.Save(inputStream, ImageFormat.Png);
            inputStream.Seek(0, SeekOrigin.Begin);
            MemoryStream outputStream = new MemoryStream();
            int size = image.Size.Width;

            image.Dispose();
            if (!ConvertToIcon(inputStream, outputStream, size, preserveAspectRatio))
            {
                return null;
            }
            return outputStream.ToArray();
        }
        public static Icon BytesToIcon(byte[] bytes)
        {
            using (MemoryStream ms = new MemoryStream(bytes))
            {
                return new Icon(ms);
            }
        }

        /// <summary>
        /// Takes a screenshot in .png format, then converts it to .ico
        /// </summary>
        /// <param name="fileName">Name of the file</param>
        /// <returns>ico byte array / null for error</returns>
        private void takeScreenshot(string fileName)
        {
            // cubic resolutions work best e.g.: (512, 512)(1024, 1024)(256, 256)
            // non-cubic resolution makes icon drawn
            var size = new Size(1024, 1024);

            // bmpScreenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, PixelFormat.Format32bppArgb);
            bmpScreenshot = new Bitmap(size.Width, size.Height, PixelFormat.Format32bppArgb);
            // Create a graphics object from the bitmap
            gfxScreenshot = Graphics.FromImage(bmpScreenshot);
            // Take the screenshot from the upper left corner to the right bottom corner
            // gfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
            var location = this.Location;
            location.Y += 80;
            location.X += 384;

            gfxScreenshot.CopyFromScreen(location, new Point(0, 0), size);
            // Save the screenshot to the specified path that the user has chosen
            bmpScreenshot.Save(fileName + ".png", ImageFormat.Png);

            bmpScreenshot.Dispose();

            string currentDir = Directory.GetCurrentDirectory();

            ConvertToIcon(
                currentDir + "\\" + fileName + ".png",
                currentDir + "\\" + fileName + "Icon" + ".png",
                512
            );

            var icon1 = ConvertToIcon(Image.FromFile(currentDir + "\\" + fileName + "Icon" + ".png"));

            using (FileStream fs = File.OpenWrite(currentDir + "\\" + currentDirectoryName + "\\" + fileName + ".ico"))
            {
                Icon ico = BytesToIcon(icon1);
                ico.Save(fs);
                ico.Dispose();
            }

            // files used to create the icon
            File.Delete(currentDir + "\\" + fileName + ".png");
            File.Delete(currentDir + "\\" + fileName + "Icon" + ".png");
        }

        /// <summary>
        /// Moves directories from x folder to y
        /// This is very important, because you can't assign icons to an existing folder
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            richTextBox1.Show();
            string[] pathList = System.IO.Directory.GetDirectories(sourceDir, "*", System.IO.SearchOption.AllDirectories);

            if (pathList.Length > 0)
            {
                richTextBox1.Text += "[START] Copying files\n\n";
                foreach (string f in pathList)
                {
                    for (var i = 0; i < extensions.Count; i++)
                    {
                        try
                        {
                            var files = Directory.GetFiles(f, extensions[i]);

                            string[] fileNames = files[0].Split('\\');
                            string[] folderName = f.Split('\\');

                            string newDirPath = newDirectory + "\\" + folderName[folderName.Length - 1];

                            DirectoryInfo directory = Directory.CreateDirectory(newDirPath);
                            directory.Attributes |= FileAttributes.ReadOnly;

                            string filePath = Path.Combine(directory.FullName, "desktop.ini");
                            FileInfo createdFile = new FileInfo(filePath);

                            try
                            {
                                // Remove the Hidden and ReadOnly attributes so file.Create*() will succeed
                                createdFile.Attributes = FileAttributes.Normal;
                            }
                            catch (FileNotFoundException)
                            {
                                // The file does not yet exist; no extra handling needed
                            }

                            using (TextWriter tw = createdFile.CreateText())
                            {
                                tw.WriteLine("[.ShellClassInfo]");
                                tw.WriteLine("IconResource=" + iconPath + folderName[folderName.Length - 1] + "\\icon.ico");
                            }

                            createdFile.Attributes = FileAttributes.ReadOnly | FileAttributes.System | FileAttributes.Hidden;

                            File.Copy(f + "\\" + fileNames[fileNames.Length - 1], newDirPath + "\\" + fileNames[fileNames.Length - 1]);


                            richTextBox1.Text += fileNames[fileNames.Length - 1];
                            richTextBox1.Text += newDirPath + "\\" + fileNames[fileNames.Length - 1] + "\n";

                            Directory.Delete(f, true);
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
                richTextBox1.Text += "\n\n[FINISH] Copying files";
            }
        }
    }
}
