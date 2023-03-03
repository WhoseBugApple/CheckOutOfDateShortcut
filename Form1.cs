using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Shell32;
using System.Security.AccessControl;

namespace CheckOutOfDateShortcut
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonRun_Click(object sender, EventArgs e)
        {
            utils.disablePublishingTask(this);
            utils.log(this.labelLog, "Running");
            String directory = textBoxInput.Text;
            if (!utils.isValidDirectory(directory))
            {
                utils.log(this.labelLog, "Invalid Directory");
                utils.enablePublishingTask(this);
                return;
            }

            new System.Threading.Thread(() =>
            {
                // find lnk files
                string[] lnkFilePaths = Directory.GetFiles(directory, "*.lnk");

                // filter lnks that missing target
                int failedCount = 0;
                IEnumerable<string> filtered = System.Linq.Enumerable.Where(
                    lnkFilePaths,
                    lnk => {
                        try  // read target and figure out is target exist
                        {
                            // Console.WriteLine(lnk);
                            String target = utils.getLnkTarget(lnk);
                            // Console.WriteLine("    " + target);
                            String dirOnly = System.IO.Path.GetDirectoryName(target);
                            // b -> is can't find target
                            bool b = Path.IsPathRooted(target) &&
                                !target.StartsWith("::{") &&
                                (!utils.ExistFileOrDirectory(target));
                            return b;
                        }
                        catch (Exception e1)
                        {
                            failedCount++;
                            Console.WriteLine("Faile to read Target for " + lnk);
                            return false;
                        }
                    });
                string[] lnks = filtered.ToArray();
                string[] targets = new string[lnks.Length];
                for(int i=0; i<targets.Length; i++)
                {
                    targets[i] = utils.getLnkTarget(lnks[i]);
                }

                // notify winform thread to show result
                this.BeginInvoke(new MyDelegate(() =>
                {
                    if (utils.showResult(listView1, lnks, targets))
                    {
                        utils.log(this.labelLog,
                            lnks.Length + " lnks Missing Target, " +
                            lnkFilePaths.Length + " lnks Checked, " + 
                            failedCount + " lnks Failed to be Checked");
                    } else
                    {
                        utils.log(this.labelLog, "Failed to Show Result");
                    }
                    utils.enablePublishingTask(this);
                }));
            }).Start();
        }
        delegate void MyDelegate();

        private Utils utils = new Utils();
        private class Utils
        {
            public void log(Label label, string text)
            {
                label.Text = text;
            }

            public bool showResult(ListView listView, string[] lnkFilePaths, string[] lnkFileTargets)
            {
                listView.Items.Clear();
                if (lnkFilePaths.Length != lnkFileTargets.Length)
                    return false;
                for (int i=0; i<lnkFilePaths.Length; i++)
                {
                    string lnk = lnkFilePaths[i];
                    string target = lnkFileTargets[i];
                    listView.Items.Add(
                        new ListViewItem(
                            new string[]{ Path.GetFileName(lnk), target}));
                }
                return true;
            }

            public bool isValidDirectory(String path)
            {
                if (path == null) return false;
                if (!Directory.Exists(path)) return false;
                return true;
            }
            
            public bool ExistFileOrDirectory(String path)
            {
                return File.Exists(path) || Directory.Exists(path);
            }

            public void disablePublishingTask(Form1 form)
            {
                form.textBoxInput.Enabled = false;
                form.buttonRun.Enabled = false;
            }

            public void enablePublishingTask(Form1 form)
            {
                form.textBoxInput.Enabled = true;
                form.buttonRun.Enabled = true;
            }

            private GetLnkTarget myGetLnkTarget = new GetLnkTargetByResolveIt();
            public String getLnkTarget(String lnk)
            {
                return myGetLnkTarget.getLnkTarget(lnk);
            }
        }
    }

    interface GetLnkTarget
    {
        String getLnkTarget(String shortcutPath);
    }

    // according to https://github.com/libyal/liblnk/blob/main/documentation/Windows%20Shortcut%20File%20(LNK)%20format.asciidoc
    class GetLnkTargetByResolveIt : GetLnkTarget
    {
        String GetLnkTarget.getLnkTarget(String shortcutPath)
        {
            FileStream stream = null;
            BinaryReader reader = null;
            try
            {
                stream = new FileStream(shortcutPath, FileMode.Open);
                reader = new BinaryReader(stream);

                // read file header
                {
                    int size = 76;  // n byte
                    reader.ReadBytes(size);
                }

                // read link target indentifier
                {
                    // int sizeOfRemain = reader.ReadInt16();
                    int size = reader.ReadInt16() + 2;
                    reader.BaseStream.Position -= 2;
                    reader.ReadBytes(size);
                }

                // read location information
                {
                    // start index of location information
                    long startIdxOfLocationInformation =
                        reader.BaseStream.Position;
                    int locationInformationSize =
                        reader.ReadInt32();
                    int locationInformationHeaderSize =
                        reader.ReadInt32();
                    reader.ReadInt32();
                    reader.ReadInt32();
                    // offset from (start of location information)
                    int offsetToLocalPath = reader.ReadInt32();

                    // start index of local path
                    long startIdxOfLocalPath =
                        startIdxOfLocationInformation + offsetToLocalPath;

                    // read until local path
                    while (reader.BaseStream.Position < startIdxOfLocalPath)
                    {
                        reader.ReadByte();
                    }
                    // read local path
                    {
                        // reader until \0
                        LinkedList<byte> bList = new LinkedList<byte>();
                        while (reader.BaseStream.Position != reader.BaseStream.Length)
                        {
                            byte bInner = reader.ReadByte();
                            if (bInner == 0) break;
                            bList.AddLast(bInner);
                        }
                        // return
                        return Encoding.Default.GetString(bList.ToArray());
                    }
                }
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            {
                stream?.Close();  // close if not null
                reader?.Close();
            }
        }
    }

    // by calling shell32.dll
    class GetLnkTargetByShell : GetLnkTarget
    {
        Shell shell = new Shell();
        String GetLnkTarget.getLnkTarget(String shortcutPath)
        {
            string pathOnly = System.IO.Path.GetDirectoryName(shortcutPath);
            string filenameOnly = System.IO.Path.GetFileName(shortcutPath);

            Folder folder = shell.NameSpace(pathOnly);
            FolderItem folderItem = folder.ParseName(filenameOnly);
            if (folderItem != null)
            {
                Shell32.ShellLinkObject link = (Shell32.ShellLinkObject)folderItem.GetLink;
                return link.Target.Path;
            }
            return string.Empty;
        }
    }
}
