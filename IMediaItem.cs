using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using IMAPI2.Interop;


namespace BurnMedia
{
    [StructLayout(LayoutKind.Sequential)]
    public struct SHFILEINFO
    {
        public IntPtr hIcon;
        public IntPtr iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    };

    class Win32
    {
        public const uint SHGFI_ICON = 0x100;
        public const uint SHGFI_LARGEICON = 0x0; // 'Large icon
        public const uint SHGFI_SMALLICON = 0x1; // 'Small icon

        [DllImport("shell32.dll")]
        public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);
    }
	
    
    interface IMediaItem
    {
        /// <summary>
        /// Returns the full path of the file or directory
        /// </summary>
        string Path { get; }

        /// <summary>
        /// Returns the size of the file or directory to the next largest sector
        /// </summary>
        Int64 SizeOnDisc { get; }

        /// <summary>
        /// Returns the Icon of the file or directory
        /// </summary>
        System.Drawing.Icon FileIcon { get; }

        // Adds the file or directory to the directory item, usually the root.
        bool AddToFileSystem(IFsiDirectoryItem rootItem);
    }


    /// <summary>
    /// 
    /// </summary>
    class FileItem : IMediaItem
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = false, EntryPoint = "SHCreateStreamOnFileW")]
        static extern void SHCreateStreamOnFile(string fileName, uint mode, ref IStream stream);

        private const uint FILE_ATTRIBUTE_NORMAL = 0x00000080; 

        private const uint STGM_DELETEONRELEASE  = 0x04000000;
        private const uint STGM_SHARE_DENY_WRITE = 0x00000020;
        private const uint STGM_SHARE_DENY_NONE  = 0x00000040;
        private const uint STGM_READ = 0x00000000;

        private const Int64 SECTOR_SIZE = 2048;

        private Int64 m_fileLength = 0;

        public FileItem(string path)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException("The file added to FileItem was not found!",path);
            }

            filePath = path;

            FileInfo fileInfo = new FileInfo(filePath);
            displayName = fileInfo.Name;
            m_fileLength = fileInfo.Length;

            //
            // Get the File icon
            //
            SHFILEINFO shinfo = new SHFILEINFO();
            IntPtr hImg = Win32.SHGetFileInfo(filePath, 0, ref shinfo, 
                (uint)Marshal.SizeOf(shinfo), Win32.SHGFI_ICON|Win32.SHGFI_SMALLICON);

            //The icon is returned in the hIcon member of the shinfo struct
            fileIcon = System.Drawing.Icon.FromHandle(shinfo.hIcon);
        }

        /// <summary>
        /// 
        /// </summary>
        public Int64 SizeOnDisc
        {
            get
            {
                if (m_fileLength > 0)
                {
                    return ((m_fileLength / SECTOR_SIZE) + 1) * SECTOR_SIZE;
                }

                return 0;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Path
        {
            get
            {
                return filePath;
            }
        }
        private string filePath;

        /// <summary>
        /// 
        /// </summary>
        public System.Drawing.Icon FileIcon
        {
            get
            {
                return fileIcon;
            }
        }
        private System.Drawing.Icon fileIcon = null;


        /// <summary>
        /// 
        /// </summary>
        public override string ToString()
        {
            return displayName;
        }
        private string displayName;

        public bool AddToFileSystem(IFsiDirectoryItem rootItem)
        {
            try
            {
                IStream stream = null;

                SHCreateStreamOnFile(filePath, STGM_READ|STGM_SHARE_DENY_WRITE, ref stream);

                if (stream != null)
                {
                    rootItem.AddFile(displayName, stream);
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error adding file", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return false;
        }

    }



    /// <summary>
    /// 
    /// </summary>
    class DirectoryItem : IMediaItem
    {
        private List<IMediaItem> mediaItems = new List<IMediaItem>();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="directoryPath"></param>
        public DirectoryItem(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                throw new FileNotFoundException("The directory added to DirectoryItem was not found!", directoryPath);
            }

            m_directoryPath = directoryPath;
            FileInfo fileInfo = new FileInfo(m_directoryPath);
            displayName = fileInfo.Name;

            //
            // Get all the files in the directory
            //
            string[] files = Directory.GetFiles(m_directoryPath);
            foreach (string file in files)
            {
                mediaItems.Add(new FileItem(file));
            }

            //
            // Get all the subdirectories
            //
            string[] directories = Directory.GetDirectories(m_directoryPath);
            foreach (string directory in directories)
            {
                mediaItems.Add(new DirectoryItem(directory));
            }

            //
            // Get the File icon
            //
            SHFILEINFO shinfo = new SHFILEINFO();
            IntPtr hImg = Win32.SHGetFileInfo(m_directoryPath, 0, ref shinfo,
                (uint)Marshal.SizeOf(shinfo), Win32.SHGFI_ICON | Win32.SHGFI_SMALLICON);

            //The icon is returned in the hIcon member of the shinfo struct
            fileIcon = System.Drawing.Icon.FromHandle(shinfo.hIcon);
        }

        /// <summary>
        /// 
        /// </summary>
        public string Path
        {
            get
            {
                return m_directoryPath;
            }
        }
        private string m_directoryPath;

        /// <summary>
        /// 
        /// </summary>
        public override string ToString()
        {
            return displayName;
        }
        private string displayName;

        /// <summary>
        /// 
        /// </summary>
        public Int64 SizeOnDisc
        {
            get
            {
                Int64 totalSize = 0;
                foreach (IMediaItem mediaItem in mediaItems)
                {
                    totalSize += mediaItem.SizeOnDisc;
                }
                return totalSize;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public System.Drawing.Icon FileIcon
        {
            get
            {
                return fileIcon;
            }
        }
        private System.Drawing.Icon fileIcon = null;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rootItem"></param>
        /// <returns></returns>
        public bool AddToFileSystem(IFsiDirectoryItem rootItem)
        {
            try
            {
                rootItem.AddTree(m_directoryPath, true);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error adding folder",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }
    }
}
