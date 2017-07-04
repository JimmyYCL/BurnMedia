using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using IMAPI2.Interop;


namespace BurnMedia
{

    public partial class MainForm : Form
    {
        private string m_clientName = "BurnMedia";

        Int64 totalDiscSize = 0;

        private bool m_isBurning = false;
        private bool m_isFormatting = false;

        private BurnData m_burnData = new BurnData();

        public MainForm()
        {
            InitializeComponent();

            //Jimmy:20170704 detect media type during loading form
            Load += buttonDetectMedia_Click;
        }

        /// <summary>
        /// Initialize the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Load(object sender, EventArgs e)
        {
            //
            // Determine the current recording devices
            //
            MsftDiscMaster2 discMaster = new MsftDiscMaster2();
            if (!discMaster.IsSupportedEnvironment)
                return;
            foreach (string uniqueRecorderID in discMaster)
            {
                MsftDiscRecorder2 discRecorder2 = new MsftDiscRecorder2();
                discRecorder2.InitializeDiscRecorder(uniqueRecorderID);

                devicesComboBox.Items.Add(discRecorder2);
            }
            if (devicesComboBox.Items.Count > 0)
            {
                devicesComboBox.SelectedIndex = 0;
            }

            //
            // Create the volume label based on the current date
            //
            DateTime now = DateTime.Now;
            textBoxLabel.Text = now.Year + "_" + now.Month + "_" + now.Day;

            labelStatusText.Text = string.Empty;
            labelFormatStatusText.Text = string.Empty;

            UpdateCapacity();

        }

        #region Device ComboBox
        /// <summary>
        /// Selected a new device
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void devicesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("devicesComboBox_SelectedIndexChanged\n");
            IDiscRecorder2 discRecorder =
                (IDiscRecorder2)devicesComboBox.Items[devicesComboBox.SelectedIndex];

            supportedMediaLabel.Text = string.Empty;
            //mediaComboBox.Items.Clear();
            //m_isCdromSupported = false;
            //m_isDvdSupported = false;
            //m_isDualLayerDvdSupported = false;
            //m_isBluraySupported = false;

            //
            // Verify recorder is supported
            //
            IDiscFormat2Data discFormatData = new MsftDiscFormat2Data();
            if (!discFormatData.IsRecorderSupported(discRecorder))
            {
                MessageBox.Show("Recorder not supported", m_clientName,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            StringBuilder supportedMediaTypes = new StringBuilder();
            foreach (IMAPI_MEDIA_PHYSICAL_TYPE mediaType in discFormatData.SupportedMediaTypes)
            {
                if (supportedMediaTypes.Length > 0)
                    supportedMediaTypes.Append(", ");
                supportedMediaTypes.Append(GetMediaTypeString(mediaType));
            }
            supportedMediaLabel.Text = supportedMediaTypes.ToString();

            ////
            //// Add Media Selection
            ////
            //if (m_isCdromSupported)
            //{
            //    mediaComboBox.Items.Add(SIZE_CD);
            //}
            //if (m_isDvdSupported)
            //{
            //    mediaComboBox.Items.Add(SIZE_DVD);
            //}
            //if (m_isDualLayerDvdSupported)
            //{
            //    mediaComboBox.Items.Add(SIZE_DVDDL);
            //}
            //if (m_isBluraySupported)
            //{
            //    mediaComboBox.Items.Add(SIZE_BLURAY);
            //}
            //mediaComboBox.SelectedIndex = 0;
        }

        /// <summary>
        /// converts an IMAPI_MEDIA_PHYSICAL_TYPE to it's string
        /// </summary>
        /// <param name="mediaType"></param>
        /// <returns></returns>
        string GetMediaTypeString(IMAPI_MEDIA_PHYSICAL_TYPE mediaType)
        {
            switch (mediaType)
            {
                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_UNKNOWN:
                default:
                    return "Unknown Media Type";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_CDROM:
                    return "CD-ROM";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_CDR:
                    return "CD-R";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_CDRW:
                    return "CD-RW";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_DVDROM:
                    return "DVD ROM";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_DVDRAM:
                    return "DVD-RAM";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_DVDPLUSR:
                    return "DVD+R";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_DVDPLUSRW:
                    return "DVD+RW";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_DVDPLUSR_DUALLAYER:
                    return "DVD+R Dual Layer";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_DVDDASHR:
                    return "DVD-R";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_DVDDASHRW:
                    return "DVD-RW";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_DVDDASHR_DUALLAYER:
                    return "DVD-R Dual Layer";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_DISK:
                    return "random-access writes";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_DVDPLUSRW_DUALLAYER:
                    return "DVD+RW DL";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_HDDVDROM:
                    return "HD DVD-ROM";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_HDDVDR:
                    return "HD DVD-R";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_HDDVDRAM:
                    return "HD DVD-RAM";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_BDROM:
                    return "Blu-ray DVD (BD-ROM)";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_BDR:
                    return "Blu-ray media";

                case IMAPI_MEDIA_PHYSICAL_TYPE.IMAPI_MEDIA_TYPE_BDRE:
                    return "Blu-ray Rewritable media";
            }
        }

        /// <summary>
        /// Provides the display string for an IDiscRecorder2 object in the combobox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void devicesComboBox_Format(object sender, ListControlConvertEventArgs e)
        {
            IDiscRecorder2 discRecorder2 = (IDiscRecorder2)e.ListItem;
            string devicePaths = string.Empty;
            string volumePath = (string)discRecorder2.VolumePathNames.GetValue(0);
            foreach (string volPath in discRecorder2.VolumePathNames)
            {
                if (!string.IsNullOrEmpty(devicePaths))
                {
                    devicePaths += ",";
                }
                devicePaths += volumePath;
            }

            e.Value = string.Format("{0} [{1}]", devicePaths, discRecorder2.ProductId);
        }
        #endregion


        #region Media Size

        private void buttonDetectMedia_Click(object sender, EventArgs e)
        {
            IDiscRecorder2 discRecorder =
                (IDiscRecorder2)devicesComboBox.Items[devicesComboBox.SelectedIndex];

            //
            // Create and initialize the IDiscFormat2Data
            //
            MsftDiscFormat2Data discFormatData = new MsftDiscFormat2Data();
            if (!discFormatData.IsCurrentMediaSupported(discRecorder))
            {
                labelMediaType.Text = "Media not supported!";
                totalDiscSize = 0;
                return;
            }
            else
            {
                //
                // Get the media type in the recorder
                //
                discFormatData.Recorder = discRecorder;
                IMAPI_MEDIA_PHYSICAL_TYPE mediaType = discFormatData.CurrentPhysicalMediaType;
                labelMediaType.Text = GetMediaTypeString(mediaType);

                //
                // Create a file system and select the media type
                //
                MsftFileSystemImage fileSystemImage = new MsftFileSystemImage();
                fileSystemImage.ChooseImageDefaultsForMediaType(mediaType);

                //
                // See if there are other recorded sessions on the disc
                //
                if (!discFormatData.MediaHeuristicallyBlank)
                {
                    fileSystemImage.MultisessionInterfaces = discFormatData.MultisessionInterfaces;
                    fileSystemImage.ImportFileSystem();
                }

                Int64 freeMediaBlocks = fileSystemImage.FreeMediaBlocks;
                totalDiscSize = 2048 * freeMediaBlocks;
            }


            UpdateCapacity();
        }

        /// <summary>
        /// Updates the capacity progressbar
        /// </summary>
        private void UpdateCapacity()
        {
            //
            // Get the text for the Max Size
            //
            if (totalDiscSize == 0)
            {
                labelTotalSize.Text = "0MB";
                return;
            }
            else if (totalDiscSize < 1000000000)
            {
                labelTotalSize.Text = string.Format("{0}MB", totalDiscSize / 1000000);
            }
            else
            {
                labelTotalSize.Text = string.Format("{0:F2}GB", (float)totalDiscSize / 1000000000.0);
            }

            //
            // Calculate the size of the files
            //
            Int64 totalMediaSize = 0;
            foreach (IMediaItem mediaItem in listBoxFiles.Items)
            {
                totalMediaSize += mediaItem.SizeOnDisc;
            }

            if (totalMediaSize == 0)
            {
                progressBarCapacity.Value = 0;
                progressBarCapacity.ForeColor = SystemColors.Highlight;
            }
            else
            {
                int percent = (int)((totalMediaSize * 100) / totalDiscSize);
                if (percent > 100)
                {
                    progressBarCapacity.Value = 100;
                    progressBarCapacity.ForeColor = Color.Red;
                }
                else
                {
                    progressBarCapacity.Value = percent;
                    progressBarCapacity.ForeColor = SystemColors.Highlight;
                }
            }
        }

        #endregion


        #region Burn Media Process

        /// <summary>
        /// User clicked the "Burn" button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonBurn_Click(object sender, EventArgs e)
        {
            if (m_isBurning)
            {
                buttonBurn.Enabled = false;
                backgroundBurnWorker.CancelAsync();
            }
            else
            {
                m_isBurning = true;
                EnableBurnUI(false);

                //m_burnData.mainForm = this;
                IDiscRecorder2 discRecorder =
                    (IDiscRecorder2)devicesComboBox.Items[devicesComboBox.SelectedIndex];
                m_burnData.uniqueRecorderId = discRecorder.ActiveDiscRecorder;

                backgroundBurnWorker.RunWorkerAsync(m_burnData);
            }
        }

        /// <summary>
        /// The thread that does the burning of the media
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundBurnWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            //
            // Create and initialize the IDiscRecorder2 object
            //
            MsftDiscRecorder2 discRecorder = new MsftDiscRecorder2();
            BurnData burnData = (BurnData)e.Argument;
            discRecorder.InitializeDiscRecorder(burnData.uniqueRecorderId);
            discRecorder.AcquireExclusiveAccess(true, m_clientName);

            //
            // Create and initialize the IDiscFormat2Data
            //
            MsftDiscFormat2Data discFormatData = new MsftDiscFormat2Data();
            discFormatData.Recorder = discRecorder;
            discFormatData.ClientName = m_clientName;
            discFormatData.ForceMediaToBeClosed = checkBoxCloseMedia.Checked;

            //
            // Check if media is blank, (for RW media)
            //
            object[] multisessionInterfaces = null;
            if (!discFormatData.MediaHeuristicallyBlank)
            {
                multisessionInterfaces = discFormatData.MultisessionInterfaces;
            }

            //
            // Create the file system
            //
            IStream fileSystem = null;
            if (!CreateMediaFileSystem(discRecorder, multisessionInterfaces, out fileSystem))
            {
                e.Result = -1;
                return;
            }

            //
            // add the Update event handler
            //
            discFormatData.Update += new DiscFormat2Data_EventHandler(discFormatData_Update);

            //
            // Write the data here
            //
            try
            {
                discFormatData.Write(fileSystem);
                e.Result = 0;
            }
            catch (COMException ex)
            {
                e.Result = ex.ErrorCode;
                MessageBox.Show(ex.Message, "IDiscFormat2Data.Write failed",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

            //
            // remove the Update event handler
            //
            discFormatData.Update -= new DiscFormat2Data_EventHandler(discFormatData_Update);

            if (this.checkBoxEject.Checked)
            {
                discRecorder.EjectMedia();
            }

            discRecorder.ReleaseExclusiveAccess();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="progress"></param>
        void discFormatData_Update([In, MarshalAs(UnmanagedType.IDispatch)] object sender, [In, MarshalAs(UnmanagedType.IDispatch)] object progress)
        {
            //
            // Check if we've cancelled
            //
            if (backgroundBurnWorker.CancellationPending)
            {
                IDiscFormat2Data format2Data = (IDiscFormat2Data)sender;
                format2Data.CancelWrite();
                return;
            }

            IDiscFormat2DataEventArgs eventArgs = (IDiscFormat2DataEventArgs)progress;

            m_burnData.task = BURN_MEDIA_TASK.BURN_MEDIA_TASK_WRITING;

            // IDiscFormat2DataEventArgs Interface
            m_burnData.elapsedTime = eventArgs.ElapsedTime;
            m_burnData.remainingTime = eventArgs.RemainingTime;
            m_burnData.totalTime = eventArgs.TotalTime;

            // IWriteEngine2EventArgs Interface
            m_burnData.currentAction = eventArgs.CurrentAction;
            m_burnData.startLba = eventArgs.StartLba;
            m_burnData.sectorCount = eventArgs.SectorCount;
            m_burnData.lastReadLba = eventArgs.LastReadLba;
            m_burnData.lastWrittenLba = eventArgs.LastWrittenLba;
            m_burnData.totalSystemBuffer = eventArgs.TotalSystemBuffer;
            m_burnData.usedSystemBuffer = eventArgs.UsedSystemBuffer;
            m_burnData.freeSystemBuffer = eventArgs.FreeSystemBuffer;

            //
            // Report back to the UI
            //
            backgroundBurnWorker.ReportProgress(0, m_burnData);
        }

        /// <summary>
        /// Completed the "Burn" thread
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundBurnWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if ((int)e.Result == 0)
            {
                labelStatusText.Text = "Finished Burning Disc!";
            }
            else
            {
                labelStatusText.Text = "Error Burning Disc!";
            }
            statusProgressBar.Value = 0;

            m_isBurning = false;
            EnableBurnUI(true);
            buttonBurn.Enabled = true;
        }

        /// <summary>
        /// Enables/Disables the "Burn" User Interface
        /// </summary>
        /// <param name="enable"></param>
        void EnableBurnUI(bool enable)
        {
            buttonBurn.Text = enable ? "&Burn" : "&Cancel";
            buttonDetectMedia.Enabled = enable;

            devicesComboBox.Enabled = enable;
            listBoxFiles.Enabled = enable;

            buttonAddFiles.Enabled = enable;
            buttonAddFolders.Enabled = enable;
            buttonRemoveFiles.Enabled = enable;
            checkBoxEject.Enabled = enable;
            checkBoxCloseMedia.Enabled = enable;
            textBoxLabel.Enabled = enable;
        }

        /// <summary>
        /// Event receives notification from the Burn thread of an event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundBurnWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //int percent = e.ProgressPercentage;
            BurnData burnData = (BurnData)e.UserState;

            if (burnData.task == BURN_MEDIA_TASK.BURN_MEDIA_TASK_FILE_SYSTEM)
            {
                labelStatusText.Text = burnData.statusMessage;
            }
            else if (burnData.task == BURN_MEDIA_TASK.BURN_MEDIA_TASK_WRITING)
            {
                switch (burnData.currentAction)
                {
                    case IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_VALIDATING_MEDIA:
                        labelStatusText.Text = "Validating current media...";
                        break;

                    case IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_FORMATTING_MEDIA:
                        labelStatusText.Text = "Formatting media...";
                        break;

                    case IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_INITIALIZING_HARDWARE:
                        labelStatusText.Text = "Initializing hardware...";
                        break;

                    case IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_CALIBRATING_POWER:
                        labelStatusText.Text = "Optimizing laser intensity...";
                        break;

                    case IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_WRITING_DATA:
                        long writtenSectors = burnData.lastWrittenLba - burnData.startLba;

                        if (writtenSectors > 0 && burnData.sectorCount > 0)
	                    {
                            int percent = (int)((100 * writtenSectors) / burnData.sectorCount);
                            labelStatusText.Text = string.Format("Progress: {0}%", percent);
                            statusProgressBar.Value = percent;
                        }
	                    else
	                    {
                            labelStatusText.Text = "Progress 0%";
                            statusProgressBar.Value = 0;
                        }
                        break;

                    case IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_FINALIZATION:
                        labelStatusText.Text = "Finalizing writing...";
                        break;

                    case IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_COMPLETED:
                        labelStatusText.Text = "Completed!";
                        break;
                }
            }
        }

        /// <summary>
        /// Enable the Burn Button if items in the file listbox
        /// </summary>
        private void EnableBurnButton()
        {
            buttonBurn.Enabled = (listBoxFiles.Items.Count > 0);
        }


        #endregion


        #region File System Process
        private bool CreateMediaFileSystem(IDiscRecorder2 discRecorder, object[] multisessionInterfaces, out IStream dataStream)
        {
            MsftFileSystemImage fileSystemImage = new MsftFileSystemImage();

            fileSystemImage.ChooseImageDefaults(discRecorder);
            fileSystemImage.FileSystemsToCreate = 
                FsiFileSystems.FsiFileSystemJoliet | FsiFileSystems.FsiFileSystemISO9660;
            fileSystemImage.VolumeName = textBoxLabel.Text;

            fileSystemImage.Update += new DFileSystemImage_EventHandler(fileSystemImage_Update);

            //
            // If multisessions, then import previous sessions
            //
            if (multisessionInterfaces != null)
            {
                fileSystemImage.MultisessionInterfaces = multisessionInterfaces;
                fileSystemImage.ImportFileSystem();
            }

	        //
	        // Get the image root
	        //
            IFsiDirectoryItem rootItem = fileSystemImage.Root;

            //
	        // Add Files and Directories to File System Image
	        //
            foreach (IMediaItem mediaItem in listBoxFiles.Items)
            {
                //
                // Check if we've cancelled
                //
                if (backgroundBurnWorker.CancellationPending)
                {
                    break;
                }

                //
                // Add to File System
                //
                mediaItem.AddToFileSystem(rootItem);
            }

            fileSystemImage.Update -= new DFileSystemImage_EventHandler(fileSystemImage_Update);

            //
            // did we cancel?
            //
            if (backgroundBurnWorker.CancellationPending)
            {
                dataStream = null;
                return false;
            }

            dataStream = fileSystemImage.CreateResultImage().ImageStream;

	        return true;
        }

        /// <summary>
        /// Event Handler for File System Progress Updates
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="currentFile"></param>
        /// <param name="copiedSectors"></param>
        /// <param name="totalSectors"></param>
        void fileSystemImage_Update([In, MarshalAs(UnmanagedType.IDispatch)] object sender,
            [In, MarshalAs(UnmanagedType.BStr)]string currentFile, [In] int copiedSectors, [In] int totalSectors)
        {
            int percentProgress = 0;
            if (copiedSectors > 0 && totalSectors > 0)
            {
                percentProgress = (copiedSectors * 100) / totalSectors;
            }

            if (!string.IsNullOrEmpty(currentFile))
            {
                FileInfo fileInfo = new FileInfo(currentFile);
                m_burnData.statusMessage = "Adding \"" + fileInfo.Name + "\" to image...";

                //
                // report back to the ui
                //
                m_burnData.task = BURN_MEDIA_TASK.BURN_MEDIA_TASK_FILE_SYSTEM;
                backgroundBurnWorker.ReportProgress(percentProgress, m_burnData);
            }

        }
        #endregion


        #region Add/Remove File(s)/Folder(s)

        /// <summary>
        /// Adds a file to the burn list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonAddFiles_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                FileItem fileItem = new FileItem(openFileDialog.FileName);
                listBoxFiles.Items.Add(fileItem);

                UpdateCapacity();
                EnableBurnButton();
            }
        }

        /// <summary>
        /// Adds a folder to the burn list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonAddFolders_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog(this) == DialogResult.OK)
            {
                DirectoryItem directoryItem = new DirectoryItem(folderBrowserDialog.SelectedPath);
                listBoxFiles.Items.Add(directoryItem);

                UpdateCapacity();
                EnableBurnButton();
            }
        }

        /// <summary>
        /// User wants to remove a file or folder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonRemoveFiles_Click(object sender, EventArgs e)
        {
            IMediaItem mediaItem = (IMediaItem)listBoxFiles.SelectedItem;
            if (mediaItem == null)
                return;

            if (MessageBox.Show("Are you sure you want to remove \"" + mediaItem + "\"?",
                "Remove item", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                listBoxFiles.Items.Remove(mediaItem);

                EnableBurnButton();
                UpdateCapacity();
            }
        }

        #endregion


        #region File ListBox Events
        /// <summary>
        /// The user has selected a file or folder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBoxFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            buttonRemoveFiles.Enabled = (listBoxFiles.SelectedIndex != -1);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBoxFiles_DrawItem(object sender, DrawItemEventArgs e)
        {
            IMediaItem mediaItem = (IMediaItem)listBoxFiles.Items[e.Index];
	        if (mediaItem == null)
            {
		        return;
            }

            e.DrawBackground();

	        if ((e.State & DrawItemState.Focus) != 0)
            {
                e.DrawFocusRectangle();
            }

            e.Graphics.DrawIcon(mediaItem.FileIcon, 2, e.Bounds.Y+2);

            RectangleF rectF = new RectangleF(e.Bounds.X + 24, e.Bounds.Y,
                e.Bounds.Width - 24, e.Bounds.Height);

            Font font = new Font(FontFamily.GenericSansSerif, 8);

            StringFormat stringFormat = new StringFormat();
            stringFormat.LineAlignment = StringAlignment.Center;
            stringFormat.Alignment = StringAlignment.Near;
            stringFormat.Trimming = StringTrimming.EllipsisCharacter;

            e.Graphics.DrawString(mediaItem.ToString(), font, new SolidBrush(e.ForeColor),
                rectF, stringFormat);
        }
        #endregion


        #region Format/Erase the Disc
        /// <summary>
        /// The user has clicked the "Format" button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonFormat_Click(object sender, EventArgs e)
        {
            m_isFormatting = true;
            EnableFormatUI(false);

            IDiscRecorder2 discRecorder =
                (IDiscRecorder2)devicesComboBox.Items[devicesComboBox.SelectedIndex];
            backgroundFormatWorker.RunWorkerAsync(discRecorder.ActiveDiscRecorder);
        }

        /// <summary>
        /// Enables/Disables the "Burn" User Interface
        /// </summary>
        /// <param name="enable"></param>
        void EnableFormatUI(bool enable)
        {
            buttonFormat.Enabled = enable;
            checkBoxEjectFormat.Enabled = enable;
            checkBoxQuickFormat.Enabled = enable;
        }

        /// <summary>
        /// Worker thread that Formats the Disc
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundFormatWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            //
            // Create and initialize the IDiscRecorder2
            //
            MsftDiscRecorder2 discRecorder = new MsftDiscRecorder2();
            string activeDiscRecorder = (string)e.Argument;
            discRecorder.InitializeDiscRecorder(activeDiscRecorder);
            discRecorder.AcquireExclusiveAccess(true, m_clientName);

            //
            // Create the IDiscFormat2Erase and set properties
            //
            MsftDiscFormat2Erase discFormatErase = new MsftDiscFormat2Erase();
            discFormatErase.Recorder = discRecorder;
            discFormatErase.ClientName = m_clientName;
            discFormatErase.FullErase = !checkBoxQuickFormat.Checked;

            //
            // Setup the Update progress event handler
            //
            discFormatErase.Update += new DiscFormat2Erase_EventHandler(discFormatErase_Update);

            //
            // Erase the media here
            //
            try
            {
                discFormatErase.EraseMedia();
                e.Result = 0;
            }
            catch (COMException ex)
            {
                e.Result = ex.ErrorCode;
                MessageBox.Show(ex.Message, "IDiscFormat2.EraseMedia failed",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

            //
            // Remove the Update progress event handler
            //
            discFormatErase.Update -= new DiscFormat2Erase_EventHandler(discFormatErase_Update);

            if (checkBoxEjectFormat.Checked)
            {
                discRecorder.EjectMedia();
            }

            discRecorder.ReleaseExclusiveAccess();
        }

        /// <summary>
        /// Event Handler for the Erase Progress Updates
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="elapsedSeconds"></param>
        /// <param name="estimatedTotalSeconds"></param>
        void discFormatErase_Update([In, MarshalAs(UnmanagedType.IDispatch)] object sender, int elapsedSeconds, int estimatedTotalSeconds)
        {
            IDiscFormat2Erase discFormat2Data = (IDiscFormat2Erase)sender;
            int percent = elapsedSeconds * 100 / estimatedTotalSeconds;
            //
            // Report back to the UI
            //
            backgroundFormatWorker.ReportProgress(percent);
        }

        private void backgroundFormatWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            labelFormatStatusText.Text = string.Format("Formatting {0}%...", e.ProgressPercentage);
            formatProgressBar.Value = e.ProgressPercentage;
        }

        private void backgroundFormatWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if ((int)e.Result == 0)
            {
                labelFormatStatusText.Text = "Finished Formatting Disc!";
            }
            else
            {
                labelFormatStatusText.Text = "Error Formatting Disc!";
            }

            formatProgressBar.Value = 0;

            m_isFormatting = false;
            EnableFormatUI(true);
        }
        #endregion

        /// <summary>
        /// Called when user selects a new tab
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            //
            // Prevent page from changing if we're burning or formatting.
            //
            if (m_isBurning || m_isFormatting)
            {
                e.Cancel = true;
            }
        }

        private void checkBoxEjectFormat_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBoxQuickFormat_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void formatProgressBar_Click(object sender, EventArgs e)
        {

        }

        private void statusProgressBar_Click(object sender, EventArgs e)
        {

        }

        private void progressBarCapacity_Click(object sender, EventArgs e)
        {

        }
    }

    public enum BURN_MEDIA_TASK
    {
        BURN_MEDIA_TASK_FILE_SYSTEM,
        BURN_MEDIA_TASK_WRITING
    }

    public class BurnData
    {
        //public MainForm mainForm;
        public string uniqueRecorderId;
        public string statusMessage;
        public BURN_MEDIA_TASK task;

        // IDiscFormat2DataEventArgs Interface
        public long elapsedTime;		// Elapsed time in seconds
        public long remainingTime;		// Remaining time in seconds
        public long totalTime;			// total estimated time in seconds
        // IWriteEngine2EventArgs Interface
        public IMAPI_FORMAT2_DATA_WRITE_ACTION currentAction;
        public long startLba;			// the starting lba of the current operation
        public long sectorCount;		// the total sectors to write in the current operation
        public long lastReadLba;		// the last read lba address
        public long lastWrittenLba;	// the last written lba address
        public long totalSystemBuffer;	// total size of the system buffer
        public long usedSystemBuffer;	// size of used system buffer
        public long freeSystemBuffer;	// size of the free system buffer
    }
}
