using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using iTunesLib;
using System.Threading;
using System.IO;

namespace iTunesHelper
{
    public partial class Form1 : Form
    {
        int numTracks;
        int totalTracks;
        string FileRoot = @"E:\mp3\";
        Dictionary<string, WorkingTrack> collection;
        List<string> files;

        int OverallTotal;
        int OverallPosition;

        WorkingTrack ActiveWorkingTrack;

        string LocationCalculation = @"{0}{1}\{2} {3}{4}\{5:00} {6} - {7}.mp3";

        DateTime dtStarted;

        iTunesAppClass iTunes;
        IITLibraryPlaylist mainLibrary;
        IITTrackCollection tracks;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dtStarted = DateTime.Now;
            backgroundWorker1.RunWorkerAsync();
        }

        private void iTunesParse()
        {
            iTunes = new iTunesAppClass();

            

            mainLibrary = iTunes.LibraryPlaylist;
            tracks = mainLibrary.Tracks;
            IITFileOrCDTrack track;

            numTracks = tracks.Count;
            totalTracks = tracks.Count;

            OverallTotal += tracks.Count;

            //files = new Dictionary<string, int>();

            TreeNode root = new TreeNode("iTunes");
            TreeNode node = root;

            while (numTracks != 0)
            {
                OverallPosition++;
                track = tracks[numTracks] as IITFileOrCDTrack;

                //track.UpdateInfoFromFile();

                if (track != null)
                {
                    if (track.Location != null)
                    {
                        if (track.Location.StartsWith(FileRoot))
                        {
                            //files.Add(track.Location, numTracks);
                            collection.Add(track.Location.ToLower(), WorkingTrackFromIITFileOrCDTrack(track, numTracks));
                        }
                    }
                }

                backgroundWorker1.ReportProgress((OverallPosition / OverallTotal) * 100);
                numTracks--;
            }

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var prog = Microsoft.WindowsAPICodePack.Taskbar.TaskbarManager.Instance;
            prog.SetProgressState(Microsoft.WindowsAPICodePack.Taskbar.TaskbarProgressBarState.Normal);

            collection = new Dictionary<string, WorkingTrack>();

            files = GetFiles(FileRoot).ToList();

            OverallTotal = files.Count;
            OverallPosition = 0;

            iTunesParse();

            FileSystemParse();

            dataGridView1.Invoke(new BindTableDelegate(BindTable));
        }

        private void FileSystemParse()
        {
            int x = 0;

            totalTracks = files.Count;

            foreach (string file in files)
            {
                if (!collection.ContainsKey(file.ToLower()))
                {
                    //WorkingTrack wtDiskOnly = new WorkingTrack();
                    //wtDiskOnly.iTunesLocation = file;
                    //wtDiskOnly.isIniTunes = false;
                    //wtDiskOnly.isOnDisk = true;

                    try
                    {
                        TagLib.File mp3file = TagLib.File.Create(file);
                        WorkingTrack wtDiskOnly = WorkingTrackFromTagLibFile(mp3file);
                        collection.Add(file.ToLower(), wtDiskOnly);
                    }
                    catch
                    {
                        System.Diagnostics.Debug.WriteLine(file + " <-- ERROR");
                    }
                }

                x++;

                OverallPosition++;

                int progress = (int)Math.Ceiling(((float)OverallPosition / (float)OverallTotal) * 100);

                //System.Diagnostics.Debug.WriteLine(progress.ToString());

                //backgroundWorker1.ReportProgress(progress);
                //numTracks--;
            }
        }

        static IEnumerable<string> GetFiles(string path)
        {
            Queue<string> queue = new Queue<string>();
            queue.Enqueue(path);
            while (queue.Count > 0)
            {
                path = queue.Dequeue();
                try
                {
                    foreach (string subDir in Directory.GetDirectories(path))
                    {
                        queue.Enqueue(subDir);
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(ex);
                }
                string[] files = null;
                try
                {
                    files = Directory.GetFiles(path, "*.mp3");
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(ex);
                }
                if (files != null)
                {
                    for (int i = 0; i < files.Length; i++)
                    {
                        yield return files[i];
                    }
                }
            }
        }

        public delegate void BindTableDelegate();

        public void BindTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("filename");
            dt.Columns.Add("itunes");
            dt.Columns.Add("disk");
            dt.Columns.Add("mismatch");
            dt.Columns.Add("approx");
            dt.Columns.Add("propercomp");

            foreach (KeyValuePair<string, WorkingTrack> track in collection)
            {
                WorkingTrack wt = track.Value;

                DataRow dr = dt.NewRow();
                dr["filename"] = wt.Path;
                dr["disk"] = wt.isOnDisk ? "X" : "";
                dr["iTunes"] = wt.isIniTunes ? "X" : "";
                dr["mismatch"] = wt.LocationMismatch ? "X" : "";
                dr["approx"] = wt.ApproximateMatch ? "X" : "";
                dr["propercomp"] = wt.ProperlyMarkedAsCompilation ? "X" : "";

                dt.Rows.Add(dr);
            }

            dataGridView1.DataSource = dt;
            dt.DefaultView.Sort = "filename asc";
            ResizeColumns();
        }

        private void ResizeColumns()
        {
            if (dataGridView1.Columns.Count > 0)
            {
                dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;

                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    int colw = dataGridView1.Columns[i].Width;
                    dataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    dataGridView1.Columns[i].Width = colw;
                }
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = Convert.ToInt32(((float)OverallPosition / (float)OverallTotal) * 100);

            TimeSpan tsElapsed = DateTime.Now - dtStarted;
            //long iSecondsLeft = Convert.ToInt64(((float)tsElapsed.Ticks / (totalTracks - numTracks)) * numTracks);
            //TimeSpan tsLeft = new TimeSpan(iSecondsLeft);

            //lblProgress.Text = string.Format("{0:n0}/{1:n0} ({2}%)", totalTracks - numTracks, totalTracks, (totalTracks - numTracks) * 100 / totalTracks);
            lblProgress.Text = string.Format("{0:n0}/{1:n0} ({2}%)", OverallPosition, OverallTotal, OverallPosition * 100 / OverallTotal);
            //lblTimeRemaining.Text = string.Format("{0:mm\\:ss}", tsLeft);

            var prog = Microsoft.WindowsAPICodePack.Taskbar.TaskbarManager.Instance;
            prog.SetProgressValue(OverallPosition, OverallTotal);
            if (OverallPosition == OverallTotal)
            {
                lblProgress.Text = "";
                //lblTimeRemaining.Text = "";
                prog.SetProgressState(Microsoft.WindowsAPICodePack.Taskbar.TaskbarProgressBarState.NoProgress);
            }
        }

        private void FieldsFromWorkingTrack(WorkingTrack WorkingTrack)
        {
            this.Message.Text = "";

            this.iTunesArtist.Text = WorkingTrack.iTunesArtist;
            this.iTunesSongName.Text = WorkingTrack.iTunesName;
            this.iTunesAlbum.Text = WorkingTrack.iTunesAlbum;
            this.iTunesAlbumArtist.Text = WorkingTrack.iTunesAlbumArtist;
            this.iTunesSortArtist.Text = WorkingTrack.iTunesSortArtist;
            this.iTunesSortName.Text = WorkingTrack.iTunesSortName;
            this.iTunesSortAlbumArtist.Text = WorkingTrack.iTunesSortAlbumArtist;
            this.iTunesTrackNumber.Text = WorkingTrack.iTunesTrackNumber;
            this.iTunesYear.Text = WorkingTrack.iTunesYear;
            this.iTunesComment.Text = WorkingTrack.iTunesComment;
            this.iTunesRating.Text = WorkingTrack.iTunesRating;

            this.iTunesFileLocation.Text = string.IsNullOrWhiteSpace(WorkingTrack.iTunesLocation) ? WorkingTrack.ID3Location : WorkingTrack.iTunesLocation;

            this.ID3Artist.Text = WorkingTrack.ID3Artist;
            this.ID3SongName.Text = WorkingTrack.ID3Name;
            this.ID3Album.Text = WorkingTrack.ID3Album;
            this.ID3AlbumArtist.Text = WorkingTrack.ID3AlbumArtist;
            this.ID3SortArtist.Text = WorkingTrack.ID3SortArtist;
            this.ID3SortName.Text = WorkingTrack.ID3SortName;
            this.ID3SortAlbumArtist.Text = WorkingTrack.ID3SortAlbumArtist;
            this.ID3TrackNumber.Text = WorkingTrack.ID3TrackNumber;
            this.ID3Year.Text = WorkingTrack.ID3Year;
            this.ID3Comment.Text = WorkingTrack.ID3Comment;
            this.ID3Rating.Text = WorkingTrack.ID3Rating;
            //this.ID3FileLocation.Text = WorkingTrack.ID3Location;
            this.iTunesCompilation.Checked = WorkingTrack.Compilation;

            if (!string.IsNullOrWhiteSpace(iTunesTrackNumber.Text))
            {
                this.iTunesCalcLocation.Text = CalculateiTunesLocation(WorkingTrack);

                if (WorkingTrack.iTunesLocation != this.iTunesCalcLocation.Text)
                {
                    this.Message.Text = "iTunes Location mismatch";

                    for (int i = 0; i < WorkingTrack.iTunesLocation.Length; i++)
                    {
                        if (this.iTunesCalcLocation.Text.Length > i)
                        {
                            char c1 = WorkingTrack.iTunesLocation[i];
                            char c2 = this.iTunesCalcLocation.Text[i];

                            if (c1 != c2)
                            {
                                this.Message.Text += string.Format(" (Starting at position {0})", i);
                                break;
                            }
                        }
                    }
                }
            }
            else
            {
                this.iTunesCalcLocation.Text = "";
            }

            if (!string.IsNullOrWhiteSpace(ID3TrackNumber.Text))
            {
                this.ID3CalcLocation.Text = CalculateID3Location(WorkingTrack);

                if (WorkingTrack.ID3Location != this.ID3CalcLocation.Text)
                {
                    this.Message.Text = "ID3 Location mismatch";

                    for (int i = 0; i < WorkingTrack.ID3Location.Length; i++)
                    {
                        if (this.ID3CalcLocation.Text.Length > i)
                        {
                            char c1 = WorkingTrack.ID3Location[i];
                            char c2 = this.ID3CalcLocation.Text[i];

                            if (c1 != c2)
                            {
                                this.Message.Text += string.Format(" (Starting at position {0})", i);
                                break;
                            }
                        }
                    }
                }
            }
            else
            {
                this.ID3CalcLocation.Text = "";
            }

            try
            {
                Artwork.Image = Image.FromStream(new MemoryStream(TagLib.File.Create(WorkingTrack.Path).Tag.Pictures[0].Data.Data));
            }
            catch
            {
                Artwork.Image = null;
            }
        }

        private string CalculateiTunesLocation(WorkingTrack WorkingTrack)
        {
            return string.Format(LocationCalculation,
                FileRoot,
                WorkingTrack.iTunesAlbumArtist,
                WorkingTrack.iTunesYear,
                WorkingTrack.iTunesAlbum,
                !string.IsNullOrWhiteSpace(WorkingTrack.iTunesComment) ? string.Format(" ({0})", WorkingTrack.iTunesComment) : "",
                int.Parse(WorkingTrack.iTunesTrackNumber),
                WorkingTrack.iTunesArtist,
                WorkingTrack.iTunesName);
        }

        private string ReCalculateiTunesLocation()
        {
            return string.Format(LocationCalculation,
                FileRoot,
                this.iTunesAlbumArtist.Text,
                this.iTunesYear.Text,
                this.iTunesAlbum.Text,
                !string.IsNullOrWhiteSpace(this.iTunesComment.Text) ? string.Format(" ({0})", this.iTunesComment.Text) : "",
                int.Parse(this.iTunesTrackNumber.Text),
                this.iTunesArtist.Text,
                this.iTunesSongName.Text);
        }

        private string CalculateID3Location(WorkingTrack WorkingTrack)
        {
            return string.Format(LocationCalculation,
                FileRoot,
                WorkingTrack.ID3AlbumArtist,
                WorkingTrack.ID3Year,
                WorkingTrack.ID3Album,
                !string.IsNullOrWhiteSpace(WorkingTrack.ID3Comment) ? string.Format(" ({0})", WorkingTrack.ID3Comment) : "",
                int.Parse(WorkingTrack.ID3TrackNumber),
                WorkingTrack.ID3Artist,
                WorkingTrack.ID3Name);
        }

        private string ReCalculateID3Location()
        {
            return string.Format(LocationCalculation,
                FileRoot,
                this.ID3AlbumArtist.Text,
                this.ID3Year.Text,
                this.ID3Album.Text,
                !string.IsNullOrWhiteSpace(this.ID3Comment.Text) ? string.Format(" ({0})", this.ID3Comment.Text) : "",
                int.Parse(this.ID3TrackNumber.Text),
                this.ID3Artist.Text,
                this.ID3SongName.Text);
        }

        private WorkingTrack MergeWorkingTrackFromFields(WorkingTrack WorkingTrack)
        {
            WorkingTrack.iTunesArtist = this.iTunesArtist.Text;
            WorkingTrack.iTunesName = this.iTunesSongName.Text;
            WorkingTrack.iTunesAlbum = this.iTunesAlbum.Text;
            WorkingTrack.iTunesAlbumArtist = this.iTunesAlbumArtist.Text;
            WorkingTrack.iTunesSortArtist = this.iTunesSortArtist.Text;
            WorkingTrack.iTunesSortName = this.iTunesSortName.Text;
            WorkingTrack.iTunesSortAlbumArtist = this.iTunesSortAlbumArtist.Text;
            WorkingTrack.iTunesTrackNumber = this.iTunesTrackNumber.Text;
            WorkingTrack.iTunesYear = this.iTunesYear.Text;
            WorkingTrack.iTunesComment = this.iTunesComment.Text;
            WorkingTrack.iTunesRating = this.iTunesRating.Text;
            WorkingTrack.iTunesLocation = this.iTunesFileLocation.Text;

            return WorkingTrack;
        }

        private WorkingTrack WorkingTrackFromIITFileOrCDTrack(IITFileOrCDTrack IITFileOrCDTrack, int Position)
        {
            WorkingTrack WorkingTrack;

            if (File.Exists(IITFileOrCDTrack.Location))
            {
                WorkingTrack = WorkingTrackFromTagLibFile(TagLib.File.Create(IITFileOrCDTrack.Location));
            }
            else
            {
                WorkingTrack = new WorkingTrack();
            }

            WorkingTrack.iTunesArtist = IITFileOrCDTrack.Artist;
            WorkingTrack.iTunesName = IITFileOrCDTrack.Name;
            WorkingTrack.iTunesAlbum = IITFileOrCDTrack.Album;
            WorkingTrack.iTunesAlbumArtist = IITFileOrCDTrack.AlbumArtist;
            WorkingTrack.iTunesSortArtist = IITFileOrCDTrack.SortArtist;
            WorkingTrack.iTunesSortName = IITFileOrCDTrack.SortName;
            WorkingTrack.iTunesSortAlbumArtist = IITFileOrCDTrack.SortAlbumArtist;
            WorkingTrack.iTunesTrackNumber = IITFileOrCDTrack.TrackNumber.ToString();
            WorkingTrack.iTunesYear = IITFileOrCDTrack.Year.ToString();
            WorkingTrack.iTunesComment = IITFileOrCDTrack.Comment;
            WorkingTrack.iTunesRating = IITFileOrCDTrack.Rating.ToString();
            WorkingTrack.iTunesLocation = IITFileOrCDTrack.Location.ToString();
            WorkingTrack.isIniTunes = true; // IITFileOrCDTrack means it's from iTunes
            WorkingTrack.isOnDisk = File.Exists(WorkingTrack.iTunesLocation);
            WorkingTrack.Position = Position;
            WorkingTrack.LocationMismatch = (WorkingTrack.iTunesLocation != CalculateiTunesLocation(WorkingTrack));

            if (WorkingTrack.LocationMismatch)
            {
                WorkingTrack.ApproximateMatch = (WorkingTrack.iTunesLocation.ToLower() == CalculateiTunesLocation(WorkingTrack).ToLower());
            }

            WorkingTrack.Compilation = IITFileOrCDTrack.Compilation;

            if (WorkingTrack.ID3Comment == "Compilation")
            {
                // If the comment says compilation you need to be flagged as compilation
                WorkingTrack.ProperlyMarkedAsCompilation = WorkingTrack.Compilation;
            }
            else
            {
                // If the comment does not say compilation you need to be NOT flagged as compilation
                WorkingTrack.ProperlyMarkedAsCompilation = !WorkingTrack.Compilation;
            }

            WorkingTrack.Path = IITFileOrCDTrack.Location;

            return WorkingTrack;
        }

        private WorkingTrack WorkingTrackFromTagLibFile(TagLib.File TagLibFile)
        {
            WorkingTrack WorkingTrack = new WorkingTrack();

            WorkingTrack.ID3Artist = TagLibFile.Tag.FirstPerformer; //??
            WorkingTrack.ID3Name = TagLibFile.Tag.Title;
            WorkingTrack.ID3Album = TagLibFile.Tag.Album;
            WorkingTrack.ID3AlbumArtist = TagLibFile.Tag.FirstAlbumArtist;
            WorkingTrack.ID3SortArtist = TagLibFile.Tag.FirstPerformerSort;
            WorkingTrack.ID3SortName = TagLibFile.Tag.TitleSort;
            WorkingTrack.ID3SortAlbumArtist = TagLibFile.Tag.FirstAlbumArtistSort;
            WorkingTrack.ID3TrackNumber = TagLibFile.Tag.Track.ToString();
            WorkingTrack.ID3Year = TagLibFile.Tag.Year.ToString();
            WorkingTrack.ID3Comment = TagLibFile.Tag.Comment;
            //WorkingTrack.iTunesRating = TagLibFile.Tag.ra.ToString();
            WorkingTrack.ID3Location = TagLibFile.Name;
            //WorkingTrack.isIniTunes = true; // IITFileOrCDTrack means it's from iTunes
            WorkingTrack.isOnDisk = File.Exists(TagLibFile.Name);
            WorkingTrack.Path = TagLibFile.Name;
            //WorkingTrack.Position = Position;
            //WorkingTrack.LocationMismatch = (WorkingTrack.ID3Location != CalculateLocation(WorkingTrack));

            //if (WorkingTrack.LocationMismatch)
            //{
            //    WorkingTrack.ApproximateMatch = (WorkingTrack.iTunesLocation.ToLower() == CalculateLocation(WorkingTrack).ToLower());
            //}

            try
            {
                //TagLib.IPicture pic = TagLibFile.Tag.Pictures[0];
                //MemoryStream stream = new MemoryStream(pic.Data.Data);
                //WorkingTrack.Artwork = Image.FromStream(stream);
            }
            catch
            {

            }

            return WorkingTrack;
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            ResizeColumns();
        }

        private void dataGridView1_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            if (e.StateChanged == DataGridViewElementStates.Selected)
            {
                string filename = e.Row.Cells[0].Value.ToString().ToLower();
                if (collection.ContainsKey(filename))
                {
                    ActiveWorkingTrack = collection[filename];
                    FieldsFromWorkingTrack(ActiveWorkingTrack);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Message.Text = "";
            lblProgress.Text = "";
            lblTimeRemaining.Text = "";

            this.WindowState = FormWindowState.Maximized;
        }

        private void chkFilteriTunes_CheckedChanged(object sender, EventArgs e)
        {
            FilterGrid();
        }

        private void FilterGrid()
        {
            string filter = "";

            if (rbiTunesYes.Checked)
            {
                filter += "itunes = 'X'";
            }
            if (rbiTunesNo.Checked)
            {
                filter += "itunes = ''";
            }

            if (rbDiskYes.Checked)
            {
                if (filter != "") filter += " or ";
                filter += "disk = 'X'";
            }
            if (rbDiskNo.Checked)
            {
                if (filter != "") filter += " or ";
                filter += "disk = ''";
            }

            if (rbMismatchYes.Checked)
            {
                if (filter != "") filter += " or ";
                filter += "mismatch = 'X'";
            }
            if (rbMismatchNo.Checked)
            {
                if (filter != "") filter += " or ";
                filter += "mismatch = ''";
            }

            if (rbApproxYes.Checked)
            {
                if (filter != "") filter += " or ";
                filter += "approx = 'X'";
            }
            if (rbApproxNo.Checked)
            {
                if (filter != "") filter += " or ";
                filter += "approx = ''";
            }

            if (rbProperCompilationYes.Checked)
            {
                if (filter != "") filter += " or ";
                filter += "propercomp = 'X'";
            }
            if (rbProperCompilationNo.Checked)
            {
                if (filter != "") filter += " or ";
                filter += "propercomp = ''";
            }

            if (dataGridView1.DataSource != null)
            {
                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = filter;
            }

            txtFilter.Text = filter;

            int a = dataGridView1.Rows.Count;
            int b = dataGridView1.Rows.GetRowCount(DataGridViewElementStates.Visible);
        }

        private void chkFilterDisk_CheckedChanged(object sender, EventArgs e)
        {
            FilterGrid();
        }

        private void chkFilterMismatch_CheckedChanged(object sender, EventArgs e)
        {
            FilterGrid();
        }

        private void chkFilterApprox_CheckedChanged(object sender, EventArgs e)
        {
            FilterGrid();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MoveFileToiTunesCalcLocation(this.iTunesFileLocation.Text);

            MessageBox.Show("done");

        }

        private string MoveFileToiTunesCalcLocation(string FileLocation)
        {
            WorkingTrack wt = collection[FileLocation.ToLower()];

            return MoveFileToiTunesCalcLocation(wt);
        }

        private string MoveFileToiTunesCalcLocation(WorkingTrack wt)
        {
            string CalcLocation = CalculateiTunesLocation(wt);

            if (!Directory.Exists(Path.GetDirectoryName(CalcLocation)))
                Directory.CreateDirectory(Path.GetDirectoryName(CalcLocation));

            File.Move(wt.iTunesLocation, CalcLocation);

            if (wt.Position != null)
            {
                IITFileOrCDTrack track = GetITTFileOrCDTrackFromWorkingTrack(wt);
                track.Location = CalcLocation;

                track.Compilation = (wt.ID3Comment == "Compilation");
            }

            return CalcLocation;
        }

        private string MoveFileToID3CalcLocation(string FileLocation)
        {
            WorkingTrack wt = collection[FileLocation.ToLower()];

            return MoveFileToID3CalcLocation(wt);
        }

        private string MoveFileToID3CalcLocation(WorkingTrack wt)
        {
            string CalcLocation = CalculateID3Location(wt);

            if (!Directory.Exists(Path.GetDirectoryName(CalcLocation)))
                Directory.CreateDirectory(Path.GetDirectoryName(CalcLocation));

            dataGridView1.Rows.GetRowCount(DataGridViewElementStates.Visible);

            File.Move(wt.ID3Location, CalcLocation);

            if (wt.Position != null)
            {
                IITFileOrCDTrack track = GetITTFileOrCDTrackFromWorkingTrack(wt);
                track.Location = CalcLocation;

                track.Compilation = (wt.ID3Comment == "Compilation");
            }

            return CalcLocation;
        }

        private IITFileOrCDTrack GetITTFileOrCDTrackFromWorkingTrack(WorkingTrack wt)
        {
            iTunes = new iTunesAppClass();
            mainLibrary = iTunes.LibraryPlaylist;
            tracks = mainLibrary.Tracks;
            IITFileOrCDTrack track = tracks[wt.Position.Value] as IITFileOrCDTrack;
            return track;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (KeyValuePair<string, WorkingTrack> track in collection)
            {
                WorkingTrack wt = track.Value;

                if (wt.ApproximateMatch)
                {
                    MoveFileToiTunesCalcLocation(wt);
                }
            }

            MessageBox.Show("done");
        }

        private void Field_TextChanged(object sender, EventArgs e)
        {
            this.iTunesCalcLocation.Text = ReCalculateiTunesLocation();
        }

        private void btnRecalc_Click(object sender, EventArgs e)
        {
            this.iTunesCalcLocation.Text = ReCalculateiTunesLocation();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            iTunes = new iTunesAppClass();
            mainLibrary = iTunes.LibraryPlaylist;
            tracks = mainLibrary.Tracks;
            IITFileOrCDTrack track = tracks[ActiveWorkingTrack.Position.Value] as IITFileOrCDTrack;

            if (track != null)
            {
                if (track.Artist != this.iTunesArtist.Text)
                    track.Artist = this.iTunesArtist.Text;

                if (track.Name != this.iTunesSongName.Text)
                    track.Name = this.iTunesSongName.Text;

                if (track.Album != this.iTunesAlbum.Text)
                    track.Album = this.iTunesAlbum.Text;

                if (track.AlbumArtist != this.iTunesAlbumArtist.Text)
                    track.AlbumArtist = this.iTunesAlbumArtist.Text;

                if (track.SortArtist != this.iTunesSortArtist.Text)
                    track.SortArtist = this.iTunesSortArtist.Text;

                if (track.SortName != this.iTunesSortName.Text)
                    track.SortName = this.iTunesSortName.Text;

                if (track.SortAlbumArtist != this.iTunesSortAlbumArtist.Text)
                    track.SortAlbumArtist = this.iTunesSortAlbumArtist.Text;

                if (track.TrackNumber.ToString() != this.iTunesTrackNumber.Text)
                    track.TrackNumber = int.Parse(this.iTunesTrackNumber.Text);

                if (track.Year.ToString() != this.iTunesYear.Text)
                    track.Year = int.Parse(this.iTunesYear.Text);

                if (track.Comment != this.iTunesComment.Text)
                    track.Comment = this.iTunesComment.Text;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                string x = row.Cells[0].Value.ToString();
                string newLocation = MoveFileToiTunesCalcLocation(x);

                collection[x.ToLower()].iTunesLocation = newLocation;
                collection[x.ToLower()].ID3Location = newLocation;
                row.Cells[0].Value = newLocation;
            }
            MessageBox.Show("done");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                string x = row.Cells[0].Value.ToString();

                WorkingTrack wt = collection[x.ToLower()];

                if (wt != null)
                {
                    if (wt.Position != null)
                    {
                        IITFileOrCDTrack track = GetITTFileOrCDTrackFromWorkingTrack(wt);

                        string AlbumArtist = new FileInfo(wt.iTunesLocation).Directory.Parent.Name;

                        if (track.AlbumArtist != AlbumArtist)
                            track.AlbumArtist = AlbumArtist;

                        wt.iTunesAlbumArtist = AlbumArtist;
                    }

                    string newID3AlbumArtist = new FileInfo(wt.ID3Location).Directory.Parent.Name;

                    TagLib.File mp3File = TagLib.File.Create(wt.ID3Location);

                    if (mp3File.Tag.FirstAlbumArtist != newID3AlbumArtist)
                    {
                        mp3File.Tag.AlbumArtists = new string[] { newID3AlbumArtist };
                        mp3File.Save();
                    }

                    wt.ID3AlbumArtist = newID3AlbumArtist;

                    collection[x.ToLower()] = wt;
                }
            }
            MessageBox.Show("done");
        }

        private void chkProperComp_CheckedChanged(object sender, EventArgs e)
        {
            FilterGrid();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = txtFilter.Text;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //System.Diagnostics.Process.Start(Path.GetDirectoryName(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString()));
            string mp3tagcmd = "\"C:\\Program Files (x86)\\Mp3tag\\Mp3tag.exe\"";
            string mp3tagparams = string.Format("/fp:\"{0}\"", Path.GetDirectoryName(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString().Replace("\\", "\\\\")));
            System.Diagnostics.Process.Start(mp3tagcmd, mp3tagparams);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                string x = row.Cells[0].Value.ToString();
                IITOperationStatus result = mainLibrary.AddFile(x);
            }
            MessageBox.Show("done");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                string x = row.Cells[0].Value.ToString();

                if (File.Exists(x))
                    RescanFile(x);
            }
            MessageBox.Show("done");
        }

        private void RescanFile(string x)
        {
            TagLib.File mp3file = TagLib.File.Create(x);
            WorkingTrack wtDiskOnly = WorkingTrackFromTagLibFile(mp3file);

            collection[x.ToLower()] = wtDiskOnly;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                string x = row.Cells[0].Value.ToString();
                string newLocation = MoveFileToID3CalcLocation(x);

                collection[x.ToLower()].iTunesLocation = newLocation;
                collection[x.ToLower()].ID3Location = newLocation;
                row.Cells[0].Value = newLocation;

                RescanFile(newLocation);
            }
            MessageBox.Show("done");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                string x = row.Cells[0].Value.ToString();

                TagLib.File mp3file = TagLib.File.Create(x);

                mp3file.Tag.Comment = "";
                mp3file.Save();

                WorkingTrack wtDiskOnly = WorkingTrackFromTagLibFile(mp3file);

                collection[x.ToLower()] = wtDiskOnly;
            }
            MessageBox.Show("done");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            foreach (WorkingTrack wt in collection.Values)
            {
                if (wt.isIniTunes)
                {
                    sb.AppendFormat("{0}\t{1}\t{2}\t{3}\t{4}", wt.iTunesArtist, wt.iTunesAlbum, wt.iTunesName, wt.iTunesLocation, wt.iTunesRating);
                    sb.Append(Environment.NewLine);
                }
            }

            File.WriteAllText("ratings.txt", sb.ToString());

            System.Diagnostics.Process.Start("ratings.txt");
        }

        private void rbiTunesYes_CheckedChanged(object sender, EventArgs e)
        {
            FilterGrid();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                string x = row.Cells[0].Value.ToString();

                WorkingTrack wt = collection[x.ToLower()];

                if (wt != null)
                {
                    if (wt.Position != null)
                    {
                        IITFileOrCDTrack track = GetITTFileOrCDTrackFromWorkingTrack(wt);

                        string SortAlbum = new FileInfo(wt.iTunesLocation).Directory.Name;

                        if (track.SortAlbum != SortAlbum)
                            track.SortAlbum = SortAlbum;

                        wt.iTunesSortAlbum = SortAlbum;
                    }

                    string newID3SortAlbum = new FileInfo(wt.ID3Location).Directory.Name;

                    TagLib.File mp3File = TagLib.File.Create(wt.ID3Location);

                    if (mp3File.Tag.AlbumSort != newID3SortAlbum)
                    {
                        mp3File.Tag.AlbumSort = newID3SortAlbum;
                        mp3File.Save();
                    }

                    wt.ID3AlbumArtist = newID3SortAlbum;

                    collection[x.ToLower()] = wt;
                }
            }
            MessageBox.Show("done");
        }
    }

    public class WorkingTrack
    {
        public string iTunesArtist;
        public string iTunesName;
        public string iTunesAlbum;
        public string iTunesAlbumArtist;
        public string iTunesSortArtist;
        public string iTunesSortName;
        public string iTunesSortAlbum;
        public string iTunesSortAlbumArtist;
        public string iTunesTrackNumber;
        public string iTunesYear;
        public string iTunesComment;
        public string iTunesRating;
        public string iTunesLocation;
        public string ID3Artist;
        public string ID3Name;
        public string ID3Album;
        public string ID3AlbumArtist;
        public string ID3SortArtist;
        public string ID3SortName;
        public string ID3SortAlbum;
        public string ID3SortAlbumArtist;
        public string ID3TrackNumber;
        public string ID3Year;
        public string ID3Comment;
        public string ID3Rating;
        public string ID3Location;
        public int? Position;
        public bool isIniTunes;
        public bool isOnDisk;
        public bool LocationMismatch;
        public bool ApproximateMatch;
        public bool Compilation;
        public bool ProperlyMarkedAsCompilation;
        public string Path;
        public Image Artwork;
    }

}
